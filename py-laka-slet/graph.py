from configparser import SectionProxy
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider)
from msgraph import GraphRequestAdapter, GraphServiceClient
from msgraph.generated.me.me_request_builder import MeRequestBuilder
from msgraph.generated.me.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder)
from msgraph.generated.me.send_mail.send_mail_post_request_body import SendMailPostRequestBody
from msgraph.generated.models.message import Message
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.email_address import EmailAddress
from async_auth import AsyncDeviceCodeCredential

class Graph:
    settings: SectionProxy
    device_code_credential: AsyncDeviceCodeCredential
    adapter: GraphRequestAdapter
    user_client: GraphServiceClient

    def __init__(self, config: SectionProxy):
        self.settings = config
        client_id = self.settings['clientId']
        tenant_id = self.settings['tenantId']
        graph_scopes = self.settings['graphUserScopes'].split(' ')

        self.device_code_credential = AsyncDeviceCodeCredential(client_id, tenant_id = tenant_id)
        auth_provider = AzureIdentityAuthenticationProvider(
            self.device_code_credential,
            scopes=graph_scopes)
        self.adapter = GraphRequestAdapter(auth_provider)
        self.user_client = GraphServiceClient(self.adapter)

    async def get_user_token(self):
        graph_scopes = self.settings['graphUserScopes']
        access_token = await self.device_code_credential.get_token(graph_scopes)
        return access_token.token

    async def get_inbox(token, person: str):
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            # Only request specific properties
            select=['from', 'isRead', 'receivedDateTime', 'subject'],
            # Get at most 25 results
            top=25,
            # Sort by received time, newest first
            orderby=['receivedDateTime DESC']
        )
        request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters= query_params
        )

        messages = await token.user_client.person().mail_folders_by_id('inbox').messages().get(
                request_configuration=request_config)
        return messages