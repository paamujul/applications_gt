import base64
import sys
import types

# Provide minimal stubs for external Google modules so that ``gmail`` can be
# imported without the real dependencies installed.  Only the attributes used
# in ``gmail.py`` are defined.
sys.modules.setdefault("googleapiclient", types.ModuleType("googleapiclient"))
sys.modules["googleapiclient.discovery"] = types.ModuleType("googleapiclient.discovery")
sys.modules["googleapiclient.discovery"].build = lambda *a, **k: None
sys.modules["googleapiclient.errors"] = types.ModuleType("googleapiclient.errors")
class _HttpError(Exception):
    pass
sys.modules["googleapiclient.errors"].HttpError = _HttpError

sys.modules.setdefault("google", types.ModuleType("google"))
sys.modules["google.oauth2"] = types.ModuleType("google.oauth2")
sys.modules["google.oauth2.credentials"] = types.ModuleType("google.oauth2.credentials")
class _Credentials:
    pass
sys.modules["google.oauth2.credentials"].Credentials = _Credentials

sys.modules["google_auth_oauthlib"] = types.ModuleType("google_auth_oauthlib")
sys.modules["google_auth_oauthlib.flow"] = types.ModuleType("google_auth_oauthlib.flow")
class _Flow:
    @staticmethod
    def from_client_secrets_file(*a, **k):
        class _F:
            def run_local_server(self, port=0):
                pass
        return _F()
sys.modules["google_auth_oauthlib.flow"].InstalledAppFlow = _Flow

from gmail import decode_part


def test_decode_part_handles_missing_padding():
    text = "Hello, world!"
    encoded = base64.urlsafe_b64encode(text.encode()).decode().rstrip("=")
    assert decode_part(encoded) == text
