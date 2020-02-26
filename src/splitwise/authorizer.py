from splitwise import Splitwise
from .group_expense_manager import GroupExpenseManager


class Authorizer:
    """
    Required to access data of splitwise user.
    """

    def __init__(
            self,
            api_key: str,
            api_secret: str,
    ):
        self._api_key = api_key
        self._api_secret = api_secret

        self.splitwise_obj = Splitwise(
            self._api_key,
            self._api_secret
        )
        self._secret = None
        self._token = None

        self._access_token = None

    def get_auth_url(self):
        """
        Beginning of an authorization. Will return url.
        After user will confirm action, verifier will be displayed.
        It must be used an input argument in authorized method
        :return:
        """
        url, secret = self.splitwise_obj.getAuthorizeURL()
        self._secret = secret
        self._token = url[url.index('oauth_token=') + len('oauth_token='):]
        return url

    def authorize(self, verifier: str):
        """
        This method must be used only after get_auth_url call.
        :param verifier: id from authorization page.
        :return:
        """
        if self._secret is None:
            raise ValueError('Secret and token must be set up first, cal get_auth_url')

        self._access_token = self.splitwise_obj.getAccessToken(
            oauth_token=self._token,
            oauth_token_secret=self._secret,
            oauth_verifier=verifier
        )
        self.splitwise_obj.setAccessToken(self._access_token)

    def create_group_expence_manager(self, group_id):
        return GroupExpenseManager(self.splitwise_obj, group_id=group_id)
