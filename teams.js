const uuid = require('uuid');
const axios = require('axios');
const querystring = require('querystring');
const crypto = require('crypto');

const tenantAuthoriseURL = (tenantId) => `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;

const tenantTokenURL = (tenantId) => `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;


/*
https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?
client_id=6731de76-14a6-49ae-97bc-6eba6914391e
&response_type=code
&redirect_uri=http%3A%2F%2Flocalhost%2Fmyapp%2F
&response_mode=query
&scope=openid%20offline_access%20https%3A%2F%2Fgraph.microsoft.com%2Fmail.read
&state=12345
*/

const scope = 'https%3A%2F%2Fgraph.microsoft.com/User.Read%20https%3A%2F%2Fgraph.microsoft.com/Group.ReadWrite.All%20https%3A%2F%2Fgraph.microsoft.com/ChannelMessage.Send%20';


const codeChallengeOf = (codeVerifier) => {
  const hash = crypto.createHash('sha256');
  hash.update(codeVerifier);
  const digest = hash.digest('base64');
  // Base64 -> Base64URL
  return digest
    .replace(/=/g, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');
}

exports.authorizeURL = () => {
  const codeVerifier = uuid.v4();
  const codeChallenge = codeChallengeOf(codeVerifier);
  const query = querystring.stringify({
    tenant: process.env.TENANT_ID,
    client_id: process.env.CLIENT_ID,
    response_type: 'code',
    redirect_uri: process.env.REDIRECT_URI,
    response_mode: 'query',
    scope,
    state: uuid.v4(),
    code_challenge: codeChallenge,
    code_challenge_method: 'S256'
  });
  return {
    codeChallenge,
    codeVerifier,
    url: `${tenantAuthoriseURL(process.env.TENANT_ID)}?${query}`
  };
};

const getAccessToken = async (code, codeVerifier) => {
  const url = tenantTokenURL(process.env.TENANT_ID);
  const params = {
    tenant: process.env.TENANT_ID,
    client_id: process.env.CLIENT_ID,
    scope,
    code,
    redirect_uri: process.env.REDIRECT_URI,
    grant_type: 'authorization_code',
    client_secret: process.env.CLIENT_SECRET_TOKEN,
    code_verifier: codeVerifier
  };

  const result = await axios.post(url, querystring.stringify(params), {
    headers: {
      'Content-Type':'application/x-www-form-urlencoded',
    }
  });

  return result;
}

exports.getAccessToken = getAccessToken;
