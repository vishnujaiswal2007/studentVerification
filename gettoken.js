import fs from 'fs';
import { google } from 'googleapis'
import open from 'open';
import readline from 'readline';

//Load Credentails
const CREDENTIAL_PATH = 'credentials.json';
const TOKEN_PATH = 'token.json'


// Define desired scopes
const SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify',
  ];


const authorize = async (credentials) =>{
    const { client_secret, client_id, redirect_uris } = credentials.web
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

    //Generate URL with all neede scopes
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type:'offline',
        scope: SCOPES,
        prompt: 'consent', //Force consent to get all scopes again
    })
    // console.log('🔑 Authorize this app by visiting this URL:\n', authUrl);
    open(authUrl);

    const rl = readline.createInterface({
        input:process.stdin,
        output:process.stdout
    })

    rl.question('\n Enter the code from the page here:', async (code)=>{
        rl.close()
        try {

            const  { tokens }  = await oAuth2Client.getToken(code);
            oAuth2Client.setCredentials(tokens);
            
            //Save the token

            fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2 ));
            console.log('\n✅ Token stored to', TOKEN_PATH)
            
        } catch (error) {
            console.error('Error retrieving access token', error);
        }
    })


}

//Run it
fs.readFile(CREDENTIAL_PATH, async (error, content)=>{
if(error) return console.log("Unable to read Credential.JSON", error)
   await authorize (JSON.parse(content))
})