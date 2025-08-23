from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from langchain_openai import AzureChatOpenAI
from langgraph.graph import StateGraph, START, END
from pydantic import BaseModel, Field
from typing import Annotated
from langchain_core.messages import AnyMessage, SystemMessage, HumanMessage
from langgraph.graph.message import add_messages
from operator import add
from docx import Document
from src.components.replace_field_text import replace_text_from_json
import json
from typing import TypedDict


# -------- AZURE AUTHENTICATION --------
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
    temperature=0.1,
)
tools = [replace_text_from_json]
llm_with_tools = llm.bind_tools(tools, tool_choice="replace_text_from_json")


class OverallState(TypedDict):
    messages: Annotated[list[AnyMessage], add_messages]
    # document: Annotated[list[bytes], add]


DEFAULT_SYSTEM_PROMPT = """Du er en hjælper, der skal hjælpe med at generere Word-fletfelter. Du vil modtage en bruger-prompt, og dernæst tekstindholdet af et dokument. Dit job er, ud fra brugerprompten, at identificere hvilke passager af dokument-teksten der skal udskiftes, og hvad de skal udskiftes med."""


def find_replacements_llm(state: OverallState):
    output = llm_with_tools.invoke(state.messages)
    return {"messages": [output]}


graph_builder = StateGraph(OverallState)

# *** NODES ***
graph_builder.add_node(
    "llm_with_tools",
    llm_with_tools,
)

# *** EDGES ***
graph_builder.add_edge(START, "llm_with_tools")
graph_builder.add_edge("llm_with_tools", END)

# memory = MemorySaver
graph = graph_builder.compile()


# def start_graph_llm(user_prompt: str, document_bytes: bytes):
def start_graph_llm(user_prompt: str):

    from docx import Document
    import io

    # doc = Document(io.BytesIO(document_bytes))
    # document_text = "\n".join([para.text for para in doc.paragraphs])
    initial_message = (
        "\n---BRUGERPROMPT:---\n"
        + user_prompt
        + "\n---DOKUMENT-TEKST---\n"
        # + document_text
    )
    # state = OverallState(
    #     messages=[
    #         SystemMessage(content=DEFAULT_SYSTEM_PROMPT),
    #         HumanMessage(content=initial_message),
    #     ],
    #     document=[document_bytes],
    # )
    messages = [HumanMessage(content=initial_message)]
    # graph.invoke({"messages": messages, "document": [document_bytes]})
    val = graph.invoke(messages)
    return


def start_graph_llm_fake(user_prompt: str, document: Document):
    class FakeLLMResponse:
        def __init__(self, content):
            self.content = content

    replacements = [
        {
            "fullText": "If betingelse Borger enlig ved ældrecheck berettigelse”din” Else ”jeres”",
            "replacementText": '{ IF "J" = "{ MERGEFIELD Borger enlig ved ældrecheck berettigelse }" " din" " jeres" }',
        },
        {
            "fullText": "If betingelse Borger enlig ved ældrecheck berettigelse  ”din” Else ”jeres”",
            "replacementText": '{ IF "J" = "{ MERGEFIELD Borger enlig ved ældrecheck berettigelse }" " din" " jeres" }',
        },
    ]
    return FakeLLMResponse(json.dumps(replacements))
