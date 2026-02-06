# Copyright (c) Microsoft. All rights reserved.
"""
Multi-Agent Demo: Assistant with OutlookAgent as a Tool

This sample demonstrates using Microsoft Agent Framework to build two agents:
1. Assistant - A general-purpose assistant built on Foundry Agent Service
2. OutlookAgent - Connects to a Copilot Studio agent that manages Outlook calendars

The Assistant can use OutlookAgent as a tool to delegate calendar-related tasks.

Prerequisites:
- `az login` (Azure CLI authentication for Foundry)
- Environment variables configured in .env file
- A Copilot Studio agent deployed for Outlook calendar management

Note: The `--pre` flag is required when installing agent-framework packages 
while the Microsoft Agent Framework is in preview:
    pip install agent-framework-azure-ai --pre
    pip install agent-framework-copilotstudio --pre
"""

import asyncio
import os

from dotenv import load_dotenv
from azure.identity import InteractiveBrowserCredential

from agent_framework import ChatAgent
from agent_framework.azure import AzureAIProjectAgentProvider
from agent_framework.microsoft import CopilotStudioAgent


# Load environment variables from .env file
load_dotenv()


async def chat_loop(assistant: ChatAgent) -> None:
    """Run an interactive chat loop with the assistant.
    
    Args:
        assistant: The ChatAgent to interact with.
    """
    print("\n" + "=" * 60)
    print("Multi-Agent Assistant (with Outlook Calendar Support)")
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
    
    # Create the OutlookAgent (connects to Copilot Studio)
    outlook_agent = CopilotStudioAgent(
        name="OutlookAgent",
        description="An agent that manages Outlook calendars. Use this for scheduling meetings, "
                    "viewing calendar events, checking availability, and other calendar operations.",
    )
    print("✓ OutlookAgent initialized (Copilot Studio)")
    
    # Create a wrapper function for the OutlookAgent that collects all streaming responses
    # This is needed because the agent-as-tool pattern has issues with streaming agents
    outlook_thread = outlook_agent.get_new_thread()
    
    async def outlook_calendar_tool(request: str) -> str:
        """Query the Outlook calendar agent for calendar-related tasks.
        
        Args:
            request: The calendar request or question to process.
            
        Returns:
            The response from the Outlook calendar agent.
        """
        response_parts = []
        async for update in outlook_agent.run_stream(request, thread=outlook_thread):
            if update.text:
                response_parts.append(update.text)
        return "".join(response_parts)
    
    # Create the main Assistant with OutlookAgent as a tool
    # Use InteractiveBrowserCredential for the Azure/Foundry tenant (different from Power Platform tenant)
    credential = InteractiveBrowserCredential(tenant_id="305b5b32-6244-4e1f-bca6-058ce94a28a4")
    provider = AzureAIProjectAgentProvider(credential=credential)
    
    # Create the main assistant with the OutlookAgent as a tool
    assistant = await provider.create_agent(
        name="Assistant",
        instructions="""You are a helpful general-purpose assistant. You can help with a wide 
variety of tasks including answering questions, writing content, analyzing information, and more.

For any calendar-related requests (scheduling meetings, viewing events, checking availability, 
managing appointments), use the outlook_calendar tool to delegate to the specialized calendar agent.

Be friendly, clear, and helpful in your responses.""",
        description="A general-purpose assistant that can help with various tasks and "
                    "delegate calendar operations to the OutlookAgent.",
        tools=[outlook_calendar_tool],
    )
    print("✓ Assistant initialized (Foundry Agent Service)")
    
    return assistant


def main() -> None:
    """Main entry point for the multi-agent demo."""
    import sys
    
    # Check for --test-outlook flag to directly test the OutlookAgent
    if "--test-outlook" in sys.argv:
        asyncio.run(test_outlook_agent())
        return
    
    # Run the main chat loop
    asyncio.run(run_chat())


async def run_chat() -> None:
    """Set up agents and run the chat loop."""
    assistant = await setup_agents()
    await chat_loop(assistant)


async def test_outlook_agent() -> None:
    """Test the OutlookAgent directly with an interactive loop."""
    print("Testing OutlookAgent directly...")
    print("-" * 50)
    
    outlook_agent = CopilotStudioAgent(
        name="OutlookAgent",
        description="An agent that manages Outlook calendars.",
    )
    print("✓ OutlookAgent initialized")
    
    thread = outlook_agent.get_new_thread()
    
    print("\nType your message and press Enter to chat with OutlookAgent.")
    print("Type 'exit' or 'quit' to end.\n")
    
    while True:
        try:
            user_input = input("You: ").strip()
            
            if not user_input:
                continue
            
            if user_input.lower() in ("exit", "quit"):
                print("\nGoodbye!")
                break
            
            print("\nOutlookAgent: ", end="", flush=True)
            
            full_response = []
            try:
                async for update in outlook_agent.run_stream(user_input, thread=thread):
                    if update.text:
                        print(update.text, end="", flush=True)
                        full_response.append(update.text)
                    # Check for any error/status info
                    if hasattr(update, 'error') and update.error:
                        print(f"\n[Error: {update.error}]")
                    if hasattr(update, 'status') and update.status:
                        print(f"\n[Status: {update.status}]")
            except Exception as e:
                print(f"\n[Exception during streaming: {type(e).__name__}: {e}]")
            
            print("\n")
            print(f"  [Debug: {len(full_response)} chunks, total length: {sum(len(c) for c in full_response)} chars]")
            print()
            
        except KeyboardInterrupt:
            print("\n\nInterrupted. Goodbye!")
            break


if __name__ == "__main__":
    main()
