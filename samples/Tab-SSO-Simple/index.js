import express from "express";
import fetch from "node-fetch";
import path from "path";
const app = express();
import dotenv from 'dotenv';
dotenv.config();
const PORT = process.env.PORT;;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const appScope = process.env.APP_GRAPH_SCOPES;
//home
app.get('/', (req, res) => {
  console.log('Hello from home page');
  res.send('Hello from home page');
});
//main page (redirect from teams)
app.get("/sso", (req, res) => {
  const _dirname = path.resolve(path.dirname(''));
  res.sendFile(path.join(_dirname, ".", "sso-tab.html"));
})
//server side call for token exchange
app.get('/token', (req, res) => {
  const idToken = req.query.token;
  if (!idToken) {
    res.status(500).send("No Id Token");
    return;
  }else{
    console.log('\x1b[33m%s\x1b[0m',"Easy peasy Id Token...");
    console.log("-----------------------------------------");
    console.log(idToken);
    console.log("\x1b[32m","-----------------------------------------");
  }
  var oboPromise = new Promise((resolve, reject) => {
    const url = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
    const params = {
      "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
      "client_id": clientId,
      "client_secret": clientSecret,
      "scope": appScope,
      "requested_token_use": "on_behalf_of",
      "assertion": idToken
    };
    fetch(url, {
      method: "POST",
      body: toQueryString(params),
      headers: {
        Accept: "application/json",
        "Content-Type": "application/x-www-form-urlencoded"
      }
    }).then(result => {
      if (result.status !== 200) {
        result.json().then(json => {
          reject({ "error": json["error"] });
        });
      } else {

        result.json().then(json => {
          resolve(json);
        });
      }
    });
  });

  oboPromise.then((result) => {
      console.log('\x1b[36m%s\x1b[0m', ' Oh my heavens, it is the access token! ');
      console.log("-----------------------------------------");
      console.log(result["access_token"]);
      console.log("\x1b[32m", "-----------------------------------------");
      //graph call with the access token
      fetch("https://graph.microsoft.com/v1.0/me/",
        {
          method: 'GET',
          headers: {
            "accept": "application/json",
            "authorization": "bearer " + result["access_token"]
          },
          mode: 'cors',
          cache: 'default'
        })
        .then(res => res.json())
        .then(json => {
          res.send(json);
        });

    }, (err) => {
      console.log(err); // Error: 
      res.send(err);
    });
});
function toQueryString(queryParams) {
  let encodedQueryParams = [];
  for (let key in queryParams) {
    encodedQueryParams.push(key + "=" + encodeURIComponent(queryParams[key]));
  }
  return encodedQueryParams.join("&");
}
//start listening to server side calls
app.listen(PORT, () => {
  console.log(`Server is Running on Port ${PORT}`);
});
