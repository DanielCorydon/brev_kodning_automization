def get_token_provider_default():
    from azure.identity import DefaultAzureCredential, get_bearer_token_provider

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
    return token_provider


def get_token_provider_streamlit_secrets():
    import streamlit as st
    from azure.identity import ClientSecretCredential, get_bearer_token_provider

    tenant_id = st.secrets["AZURE_IDENTITY_TENANT_ID"]
    client_id = st.secrets["AZURE_IDENTITY_CLIENT_ID"]
    client_secret = st.secrets["AZURE_IDENTITY_CLIENT_SECRET"]
    credential = ClientSecretCredential(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
    )
    token_provider = get_bearer_token_provider(
        credential, "https://cognitiveservices.azure.com/.default"
    )
    return token_provider
