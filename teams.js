const uuid = require('uuid');
const axios = require('axios');
const querystring = require('querystring');
const crypto = require('crypto');

const tenantAuthoriseURL = (tenantId) => `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;

const tenantTokenURL = (tenantId) => `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;


const scope = 'https%3A%2F%2Fgraph.microsoft.com%2FUser.Read%20https%3A%2F%2Fgraph.microsoft.com%2FGroup.ReadWrite.All%20https%3A%2F%2Fgraph.microsoft.com%2FChannelMessage.Send%20';


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
