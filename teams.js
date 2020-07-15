const uuid = require('uuid');
const axios = require('axios');
const querystring = require('querystring');

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

const scope = 'ChannelMessage.Send%20Group.ReadWrite.All';

exports.authorizeURL = () => {
  const params = [
    `client_id=${process.env.CLIENT_ID}`,
    'response_type=code',
    `redirect_uri=${process.env.REDIRECT_URI}`,
    'response_mode=query',
    `scope=${scope}`,
    `state=${uuid.v4()}`
  ];
  return `${tenantAuthoriseURL(process.env.TENANT_ID)}?${params.join('&')}`; 
};

const getAccessToken = async (code) => {
  const url = tenantTokenURL(process.env.TENANT_ID);
  const params = {
    client_id: process.env.CLIENT_ID,
    scope,
    code,
    redirect_uri: process.env.REDIRECT_URI,
    grant_type: 'authorization_code',
    client_secret: process.env.CLIENT_SECRET_TOKEN
  };

  const result = await axios.post(url, querystring.stringify(params), {
    headers: {
      'Content-Type':'application/x-www-form-urlencoded',
    }
  });

  return result;
}

exports.getAccessToken = getAccessToken;
