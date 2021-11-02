import express from "express";
import dotenv from 'dotenv';

import { INTERACTION_REQUIRED_STATUS_TEXT } from "./constants.js";
import { getUserProfile } from './userProfileService.js';
import { getOboAccessToken } from './oboService.js';

dotenv.config();
const app = express();

// JSON middleware is needed if you want to parse request bodies
app.use(express.json());

// Web service returns the current user's profile
app.post('/userProfile', async (req, res) => {

  const tenantId = req.body.tenantId;
  const clientSideToken = req.body.clientSideToken;

  if (!clientSideToken) {
    res.status(500).send("No Id Token");
    return
  }

  try {
    const serverSideToken = await getOboAccessToken(tenantId, clientSideToken);
    const userProfile = await getUserProfile(serverSideToken);
    res.send(userProfile);
  }
  catch (error) {
    if (error === INTERACTION_REQUIRED_STATUS_TEXT) {
      // If here, Azure AD wants to interact with the user directly, so tell the
      // client side to display a pop-up auth box
      console.log('Interaction required');
      res.status(401).json({ status: 401, statusText: INTERACTION_REQUIRED_STATUS_TEXT });
    } else {
      console.log(`Error in /userProfile handling: ${error}`);
      res.status(500).json({ status: 500, statusText: error })
    }
  }

});

// Make environment values available on the client side
// NOTE: Do not pass any secret or sensitive values to the client!
app.get('/modules/env.js', (req, res) => {
  res.contentType("application/javascript");
  res.send(`
    export const env = {
      CLIENT_ID: "${process.env.CLIENT_ID}",
      SCOPES: "${process.env.SCOPES}",
      INTERACTION_REQUIRED_STATUS_TEXT: "${INTERACTION_REQUIRED_STATUS_TEXT}"
    };
  `);
});

// Serve static pages from /client
app.use(express.static('client'));

//start listening to server side calls
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Server is Running on Port ${PORT}`);
});
