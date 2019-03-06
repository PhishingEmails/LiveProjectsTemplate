async function getAccessToken(cookies, res) {
    // Do we have an access token cached?
    let token = cookies.graph_access_token;
  
    if (token) {
      // We have a token, but is it expired?
      // Expire 5 minutes early to account for clock differences
      const FIVE_MINUTES = 300000;
      const expiration = new Date(parseFloat(cookies.graph_token_expires - FIVE_MINUTES));
      if (expiration > new Date()) {
        // Token is still good, just return it
        return token;
      }
    }
  
    // Either no token or it's expired, do we have a
    // refresh token?
    const refresh_token = cookies.graph_refresh_token;
    if (refresh_token) {
      const newToken = await oauth2.accessToken.create({refresh_token: refresh_token}).refresh();
      saveValuesToCookie(newToken, res);
      return newToken.token.access_token;
    }
  
    // Nothing in the cookies that helps, return empty
    return null;
  }
  
  exports.getAccessToken = getAccessToken;