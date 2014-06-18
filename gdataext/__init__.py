import logging
import gdata.docs.client

__all__ = ('create_client', 'search_feed_id')

logger = logging.getLogger(__name__)


def create_client(config):
    """Create a service client."""
    client = config.SERVICE(source=config.APP_NAME)
    client.http_client.debug = config.DEBUG
    try:
        client.ClientLogin(config.LOGIN, config.PASSWORD,config.APP_NAME)
    except gdata.client.BadAuthentication:
        exit('Invalid user credentials given.')
    except gdata.client.Error:
        exit('Login Error')
    return client


def search_feed_id(target_id, feed):
    """Search a feed and return the object with the given `target_id`.
    """
    for item in feed.entry:
        id_ = item.id.text.split('/')[-1]
        if id_ == target_id:
            return item