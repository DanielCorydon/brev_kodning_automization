from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from langchain_openai import AzureChatOpenAI
from langgraph.graph import StateGraph, START, END
from pydantic import BaseModel
from typing import Annotated
from langchain_core.messages import AnyMessage, SystemMessage
from langgraph.graph.message import add_messages

credential = DefaultAzureCredential(
    exclude_environment_credential=True,
    exclude_developer_cli_credential=True,
    exclude_workload_identity_credential=True,
    exclude_managed_identity_credential=True,
    exclude_visual_studio_code_credential=True,
    exclude_shared_token_cache_credential=True,
    exclude_interactive_browser_credential=True,
)
token_provider = get_bearer_token_provider(
    credential, "https://cognitiveservices.azure.com/.default"
)
llm = AzureChatOpenAI(
    azure_endpoint="https://oai02-aiserv.openai.azure.com/",
    api_version="2024-10-21",
    azure_ad_token_provider=token_provider,
    # azure_deployment="gpt-4.1-nano",
    # azure_deployment="gpt-4.1",
    azure_deployment="gpt-4o-2024-08-06",
)


class OverallState(BaseModel):
    messages: Annotated[list[AnyMessage], add_messages]


sys_msg = SystemMessage(
    content="""Du modtager nu en tekst, hvor du får følgende opgave: Du vil se et mønster hvor der står "if betingelse" (case insensitive), <en arbitrær mængde ord - lad os kalde dem MIDTERORD>, og så på et tidspunkt vil der stå ”<et antal tegn>” Else ”<et antal tegn>”. Altså 2 citationstegn med noget indeni, og så 2 sitationstegn mere, med noget andet indeni, lad os kalde dem TEKST1 og TEKST2. Dette skal du transformere til følgende: { IF "J" { MERGEFIELD <MIDTERORD>}" " TEKST1" " TEKST2" Du skal kun ændre noget i teksten ved dette specifikke mønster. Du skal udelukkende returnere den ændrede tekst, intet andet."""
)


# --- Initial LLM ---
def initial_llm(state: OverallState):
    return {"messages": [llm.invoke(state.messages)]}


graph_builder = StateGraph(OverallState)

# *** NODES ***
graph_builder.add_node(
    "initial_llm",
    initial_llm,
)

# *** EDGES ***
graph_builder.add_edge(START, "initial_llm")
graph_builder.add_edge("initial_llm", END)

# memory = MemorySaver
graph = graph_builder.compile()
