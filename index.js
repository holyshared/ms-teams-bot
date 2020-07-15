const express = require('express');
const bodyParser = require('body-parser');

const teams = require('./teams');

const app = express();

app.set("view engine", "ejs");

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json({ type: 'application/*+json' }));

result = {};

app.get('/', (req, res) => {
  result = teams.authorizeURL();
  res.render('index.ejs', {
    url: result.url
  });
});

app.get('/callback', async (req, res) => {
  const code = req.query.code;
console.log('code------------------');
console.log(code);
  teams.getAccessToken(code, result.codeVerifier).then(result => {
    console.log(result);
    res.status(200).end('Ok');
  }).catch(err => {
    console.log(err);
    res.status(500).end('oops!');
  });
});

app.get('/token', (req, res, next) => {
  console.log('GET /token');
  console.log(req.body);
  res.status(200).end('Ok');
});

app.listen(process.env.PORT || 3000);
