# Copyright (c) Microsoft. All rights reserved.
"""
Multi-Agent Demo: Assistant with Copilot Studio Agent as a Tool

This sample demonstrates using Microsoft Agent Framework to build two agents:
1. Assistant - A general-purpose assistant built on Azure OpenAI
2. CopilotStudioAgent - Connects to a Copilot Studio agent for company internal data access

The Assistant can use CopilotStudioAgent as a tool to access internal company data.

Prerequisites:
- `az login` (Azure CLI authentication for Azure OpenAI)
- Environment variables configured in .env file
- A Copilot Studio agent deployed with access to company data

Note: The `--pre` flag is required when installing agent-framework packages 
while the Microsoft Agent Framework is in preview:
    pip install agent-framework-azure-ai --pre
    pip install agent-framework-copilotstudio --pre
"""

import asyncio
import os

from dotenv import load_dotenv
from azure.identity import AzureCliCredential
from msal import PublicClientApplication

from agent_framework import ChatAgent
from agent_framework.azure import AzureOpenAIChatClient
from agent_framework.microsoft import CopilotStudioAgent
from agent_framework.devui import serve


# Load environment variables from .env file
load_dotenv()

# Global token cache to persist across calls
_token_cache = None


def acquire_token_with_device_code() -> str:
    """Acquire Power Platform token using device code flow (delegated, cached)."""
    global _token_cache
    
    client_id = os.environ["COPILOTSTUDIOAGENT__AGENTAPPID"]
    tenant_id = os.environ["COPILOTSTUDIOAGENT__TENANTID"]
    
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["https://api.powerplatform.com/.default"]
    
    # Use persistent token cache
    if _token_cache is None:
        from msal import SerializableTokenCache
        _token_cache = SerializableTokenCache()
        cache_file = os.path.join(os.path.dirname(__file__), ".token_cache.json")
        if os.path.exists(cache_file):
            with open(cache_file, "r") as f:
                _token_cache.deserialize(f.read())
    
    app = PublicClientApplication(
        client_id=client_id,
        authority=authority,
        token_cache=_token_cache,
    )
    
    # Try silent acquisition first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes=scopes, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    
    # Fall back to device code flow
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise Exception(f"Failed to initiate device flow: {flow.get('error_description', flow)}")
    
    print(f"\n  → Go to: {flow['verification_uri']}")
    print(f"  → Enter code: {flow['user_code']}\n")
    
    result = app.acquire_token_by_device_flow(flow)
    
    # Save cache
    if _token_cache.has_state_changed:
        cache_file = os.path.join(os.path.dirname(__file__), ".token_cache.json")
        with open(cache_file, "w") as f:
            f.write(_token_cache.serialize())
    
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Failed to acquire token: {result.get('error_description', result)}")



async def chat_loop(assistant: ChatAgent) -> None:
    """Run an interactive chat loop with the assistant.
    
    Args:
        assistant: The ChatAgent to interact with.
    """
    print("\n" + "=" * 60)
    print("Multi-Agent Assistant (with Copilot Studio Integration)")
    print("=" * 60)
    print("Type your message and press Enter to chat.")
    print("Type 'exit' or 'quit' to end the conversation.")
    print("-" * 60 + "\n")
    
    # Create a thread to maintain conversation history
    thread = assistant.get_new_thread()
    
    while True:
        try:
            user_input = input("You: ").strip()
            
            if not user_input:
                continue
                
            if user_input.lower() in ("exit", "quit"):
                print("\nGoodbye!")
                break
            
            print("\nAssistant: ", end="", flush=True)
            
            # Stream the response for better user experience
            async for update in assistant.run_stream(user_input, thread=thread):
                if update.text:
                    print(update.text, end="", flush=True)
            
            print("\n")
            
        except KeyboardInterrupt:
            print("\n\nInterrupted. Goodbye!")
            break


async def setup_agents():
    """Set up and return the agents for DevUI."""
    print("Initializing agents...")
    
    # Acquire token for Power Platform using device code flow (delegated, cached)
    print("  Acquiring Power Platform token...")
    pp_token = acquire_token_with_device_code()
    print("  ✓ Power Platform token acquired")
    
    # Create credential for Azure OpenAI (Foundry) - key auth is disabled
    FOUNDRY_TENANT_ID = os.environ.get("AZURE_OPENAI_TENANT_ID")
    print("  Creating Foundry credential...")
    foundry_credential = AzureCliCredential(tenant_id=FOUNDRY_TENANT_ID)
    print("  ✓ Foundry credential ready")
    
    # Create the Copilot Studio Agent for internal data access
    # Configuration is read from environment variables:
    # - COPILOTSTUDIOAGENT__ENVIRONMENTID
    # - COPILOTSTUDIOAGENT__SCHEMANAME
    # - COPILOTSTUDIOAGENT__AGENTAPPID
    # - COPILOTSTUDIOAGENT__TENANTID
    copilot_agent = CopilotStudioAgent(
        name="CopilotStudioAgent",
        description="An agent that provides access to company internal data through Copilot Studio. "
                    "Use this for accessing internal knowledge bases, company documents, and enterprise data.",
        token=pp_token,
    )
    print("✓ CopilotStudioAgent initialized")
    
    # Create a wrapper function for the CopilotStudioAgent that collects all streaming responses
    # This is needed because the agent-as-tool pattern has issues with streaming agents
    copilot_thread = copilot_agent.get_new_thread()
    
    async def copilot_studio_tool(request: str) -> str:
        """Query the Copilot Studio agent for company internal data.
        
        Args:
            request: The request or question to process.
            
        Returns:
            The response from the Copilot Studio agent.
        """
        response_parts = []
        async for update in copilot_agent.run_stream(request, thread=copilot_thread):
            if update.text:
                response_parts.append(update.text)
        return "".join(response_parts)
    
    # Create the main Assistant using AzureOpenAIChatClient with credential (key auth is disabled)
    # Reads AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_CHAT_DEPLOYMENT_NAME from .env
    assistant = AzureOpenAIChatClient(credential=foundry_credential).as_agent(
        instructions="""You are a helpful general-purpose assistant. You can help with a wide 
variety of tasks including answering questions, writing content, analyzing information, and more.

For any requests that require access to company internal data, knowledge bases, or enterprise 
information, use the copilot_studio tool to delegate to the specialized Copilot Studio agent.

Be friendly, clear, and helpful in your responses.""",
        tools=[copilot_studio_tool],
    )
    print("✓ Assistant initialized (Azure OpenAI)")
    
    return assistant


def main() -> None:
    """Main entry point for the multi-agent demo."""
    import sys
    
    # Check for --devui flag to serve in DevUI
    if "--devui" in sys.argv:
        run_devui()
        return
    
    # Run the main chat loop
    asyncio.run(run_chat())


async def run_chat() -> None:
    """Set up agents and run the chat loop."""
    assistant = await setup_agents()
    await chat_loop(assistant)


def run_devui() -> None:
    """Set up agents and serve in DevUI."""
    print("Starting DevUI...")
    assistant = asyncio.run(setup_agents())
    print("\nLaunching DevUI at http://localhost:8080")
    print("Press Ctrl+C to stop.\n")
    serve(entities=[assistant], port=8080, auto_open=True)


if __name__ == "__main__":
    main()
