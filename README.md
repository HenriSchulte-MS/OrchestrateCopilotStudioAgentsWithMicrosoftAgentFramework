# Multi-Agent Demo: Azure OpenAI + Copilot Studio

> ⚠️ **Preview Notice**: The Microsoft Agent Framework is currently in **preview** (as of February 2026). APIs and features may change without notice. Use at your own risk.

This sample demonstrates using the [Microsoft Agent Framework](https://github.com/microsoft/agent-framework) to build a multi-agent system that combines:

1. **Assistant** - A general-purpose assistant powered by Azure OpenAI
2. **Copilot Studio Agent** - Connects to a Copilot Studio agent for accessing company internal data

The Assistant uses the Copilot Studio Agent as a tool, enabling it to delegate requests for internal company data to Copilot Studio.

## Architecture

```
┌─────────────────┐         ┌─────────────────────┐         ┌─────────────────┐
│                 │         │                     │         │                 │
│  User (DevUI    │ ──────► │  Assistant Agent    │ ──────► │  Copilot Studio │
│  or Console)    │         │  (Azure OpenAI)     │  tool   │  Agent          │
│                 │ ◄────── │                     │ ◄────── │                 │
└─────────────────┘         └─────────────────────┘         └─────────────────┘
                                    │                               │
                                    ▼                               ▼
                            Azure OpenAI Service           Power Platform / 
                                                           Copilot Studio
```

## Prerequisites

- Python 3.10+
- Azure CLI (`az login`)
- An Azure OpenAI deployment
- A Copilot Studio agent
- [App Registration with Power Platform API permission](https://github.com/microsoft/agent-framework/tree/main/python/samples/02-agents/providers/copilotstudio)

## Installation

```bash
# Create virtual environment
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # Linux/Mac

# Install dependencies
pip install -r requirements.txt
```

Alternatively, install directly:

```bash
pip install python-dotenv agent-framework==1.0.0b260130
```

## Configuration

Create a `.env` file with the following variables:

```env
# Azure OpenAI Configuration
AZURE_OPENAI_ENDPOINT=https://your-endpoint.openai.azure.com/
AZURE_OPENAI_CHAT_DEPLOYMENT_NAME=gpt-4.1
AZURE_OPENAI_TENANT_ID=your-azure-tenant-id

# Copilot Studio Agent Configuration
COPILOTSTUDIOAGENT__ENVIRONMENTID=your-environment-id
COPILOTSTUDIOAGENT__SCHEMANAME=your-agent-schema-name
COPILOTSTUDIOAGENT__AGENTAPPID=your-app-registration-client-id
COPILOTSTUDIOAGENT__TENANTID=your-power-platform-tenant-id
```

### Multi-Tenant Setup

This sample supports scenarios where Azure OpenAI and Copilot Studio are in **different tenants**:

- **Azure OpenAI**: Uses `AzureCliCredential` with `AZURE_OPENAI_TENANT_ID`
- **Copilot Studio**: Uses MSAL device code flow with `COPILOTSTUDIOAGENT__TENANTID`

Run `az login` before starting the application. The first time you run, you'll be prompted to authenticate to Power Platform via device code flow (tokens are cached for subsequent runs).

## Usage

### Console Chat

```bash
python main.py
```

### DevUI (Web Interface)

```bash
python main.py --devui
```

Opens a browser to http://localhost:8080 with an interactive chat interface.

## How It Works

1. **Token Acquisition**: 
   - Azure OpenAI uses Azure CLI credentials
   - Copilot Studio uses MSAL device code flow (delegated permissions)

2. **Agent Setup**:
   - Creates a `CopilotStudioAgent` instance connected to your agent in Copilot Studio
   - Wraps it as a tool function for the main assistant
   - Creates an `AzureOpenAIChatClient` assistant with the tool

3. **Chat Loop**:
   - User messages go to the Assistant
   - When internal data is needed, the Assistant calls the Copilot Studio tool
   - Responses are streamed back to the user

## Official Resources

- **Microsoft Agent Framework**: https://github.com/microsoft/agent-framework
- **Copilot Studio Samples**: https://github.com/microsoft/agent-framework/tree/main/python/samples/02-agents/providers/copilotstudio
- **Azure OpenAI Samples**: https://github.com/microsoft/agent-framework/tree/main/python/samples/02-agents/providers/azure_openai

