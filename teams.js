const uuid = require('uuid');

const AUTHORISE_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";

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
  return `${AUTHORISE_URL}?${params.join('&')}`; 
};
