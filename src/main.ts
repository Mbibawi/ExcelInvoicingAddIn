Office.onReady((info) => {
  // Office is ready
  if (info.host === Office.HostType.Excel) {
    // Excel-specific initialization code goes here
    console.log("Excel is ready!");

    loadMsalScript();
    showForm();
  }
  });


function loadMsalScript() {
  var token;
  const script = document.createElement("script");
  script.src = "https://alcdn.msauth.net/browser/2.17.0/js/msal-browser.min.js";
  script.onload = async () => (token = await getTokenWithMSAL());
  script.onerror = () => console.error("Failed to load MSAL.js");
  document.head.appendChild(script);
};

function showForm() {
  const formHtml = `<div id="filterForm" >
    <label for="client" > Client: </label>
    <input type="text" id="client" name="client"><br><br>
    <label for="affaire" > Affaire: </label>
    <input type ="text" id ="affaire" name="affaire"><br><br>
    <label for="nature" > Nature: </label>
    <input type ="text" id ="nature" name="nature"><br><br>
    <label for="date" > Date: </label>
    <input type="date" id="date" name="date"><br><br>
    <button onclick="filter()"> Filter Table</button>
  </div>`;

  const form = document.getElementById("form")
    if(form) form.innerHTML = formHtml;
  
}

async function filter() {
  const client = document.getElementById("client") as HTMLInputElement;
  const matter = document.getElementById("affaire") as HTMLInputElement;
  const nature = document.getElementById("nature") as HTMLInputElement;
  const date = document.getElementById("date") as HTMLInputElement;
  const criteria = {
    client: getArray(client.value) || undefined,
    matter: getArray(matter.value) || undefined,
    nature: getArray(nature.value) || undefined,
    date: getArray(date.value) || undefined
  };

  filterTable(criteria);

  function getArray(value: string): string[] {
    const array = value.split(",");
    return array.filter((el) => el);
  }
}
// Filter the Excel table based on form data
async function filterTable(criteria: { client: string[]; matter: string[]; nature: string[]; date: string[] }) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItem("LivreJournal");
    table.autoFilter.clearCriteria();

    if (criteria.client.length > 0) table.columns.getItemAt(0).filter.applyValuesFilter(criteria.client);

    if (criteria.matter.length > 0) table.columns.getItemAt(1).filter.applyValuesFilter(criteria.matter);

    if (criteria.nature.length > 0) table.columns.getItemAt(2).filter.applyValuesFilter(criteria.nature);

    if (criteria.date.length > 0) table.columns.getItemAt(3).filter.applyValuesFilter(criteria.date);

    const range = table.getDataBodyRange().getVisibleView();
    range.load("values");

    await context.sync();
    //await createWordDocument(range.values);
    uploadWordDocument(range.values, criteria.client.join(",") + "_Invoice" + "THEDATE");
  });
}
/*
// Create a new Word document based on a template and populate it with filtered data
async function createWordDocument(filtered: any[][]) {
  return console.log("filtered = ", filtered);

  await Word.run(async (context) => {
    const templateUrl = "https://your-onedrive-path/template.docx";
    const newDoc = context.application.createDocument(templateUrl);
    await context.sync();

    const table = newDoc.body.tables.getFirst();
    //const filteredData = await getFilteredData();

    //filtered.forEach(el) => {
    // table.(index + 1, row);
    //});

    await context.sync();

    const saveUrl = "https://your-onedrive-path/newDocument.docx";
    await newDoc.saveAs(saveUrl);
  });
}*/

// Get filtered data from the Excel table

function getTokenWithMSAL() {
      //const clientId = "f670878d-ed8e-4020-bb82-21ba582d0d9c"; the old one
      const clientId = "157dd297-447d-4592-b2d3-76b643b97132"; //the new one
      const tenantId = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
    //const redirectUri = "https://script-lab.public.cdn.office.net";
      const redirectUri = "msal157dd297-447d-4592-b2d3-76b643b97132://auth";
  // MSAL configuration
  const msalConfig = {
    auth: {
      clientId: clientId, // Replace with your Azure AD app's client ID
      redirectUri: redirectUri
    }
  };
//@ts-ignore
  const msalInstance = new msal.PublicClientApplication(msalConfig);

  return checkExistingAuthContext();

  // Function to check existing authentication context
  async function checkExistingAuthContext(): Promise<string | undefined|void> {
    try {
      const account = msalInstance.getAllAccounts()[0];
      if (account) {
        return acquireTokenSilently(account);
      } else {
          //return loginAndGetToken();
        //openLoginWindow()
          return getOfficeToken()
        //return getTokenWithSSO('minabibawi@gmail.com')
          //return credentitalsToken()
      }
    } catch (error) {
      console.error("Error checking auth context:", error);
      return undefined;
    }
  }

  async function credentitalsToken(){
    const msalConfig = {
      auth: {
        clientId: clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        //clientSecret: clientSecret,
      }
    }
  //@ts-ignore
    const cca = new msal.application.ConfidentialClientApplication(msalConfig);

    const tokenRequest = {
      scopes:["Files.ReadWrite"],
    }

    try {
      const response = await cca.acquireTokenByClientCredential(tokenRequest);
        return response.accessToken;
    }catch (error){
      console.log('Error acquiring Token: ', error)
      return null

    }

  }

  async function getOfficeToken() {
    try {
    //@ts-ignore
    return await OfficeRuntime.auth.getAccessToken()
      
  } catch (error) {
      console.log("Error : ", error)
      
    }
    
  }

  async function getTokenWithSSO(email: string) {
    const msalConfig = {
      auth: {
        clientId: clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        redirectUri: redirectUri,
        navigateToLoginRequestUrl: true,
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: true
      }
    };
    try {
      const response = await msalInstance.ssoSilent({
        scopes: ["Files.ReadWrite.All"],
        //scopes: ["https://graph.microsoft.com/.default"],
        loginHint: email // Forces MSAL to recognize the signed-in user
      });

      console.log("Token acquired via SSO:", response.accessToken);
      return response.accessToken;
    } catch (error) {
      console.error("SSO silent authentication failed:", error);
      return null;
    }
  }

  function openLoginWindow() {
    const loginUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}&scope=https://graph.microsoft.com/.default` ;

    // Open in a new window (only works if triggered by user action)
    const authWindow = window.open(loginUrl, "_blank", "width=500,height=600");

    if (!authWindow) {
      console.error("Popup blocked! Please allow popups.");
    }
  }

  // Function to handle login and acquire token
  async function loginAndGetToken(): Promise<string | undefined> {
    try {
      const loginRequest = {
        scopes: ["Files.ReadWrite"] // OneDrive scopes
      };

      await msalInstance.loginRedirect(loginRequest);

      return handleRedirectResponse();
    } catch (error) {
      console.error("Login error:", error);
      return undefined;
    }
    // Function to handle redirect response
    async function handleRedirectResponse(): Promise<string | undefined> {
      try {
        const authResult = await msalInstance.handleRedirectPromise();
        if (authResult && authResult.accessToken) {
          console.log("Access token:", authResult.accessToken);
          return authResult.accessToken;
        }
      } catch (error) {
        console.error("Redirect handling error:", error);
      }
      return undefined;
    }
  }
  // Function to get access token silently
  async function acquireTokenSilently(account: any): Promise<string | undefined> {
    try {
      const tokenRequest = {
        scopes: ["Files.ReadWritel"], // OneDrive scopes
        account: account
      };

      const tokenResponse = await msalInstance.acquireTokenSilent(tokenRequest);
      if (tokenResponse && tokenResponse.accessToken) {
        console.log("Access token (silent):", tokenResponse.accessToken);
        return tokenResponse.accessToken;
      }
    } catch (error) {
      console.error("Token acquisition error:", error);
      return loginAndGetToken();
    }
    return undefined;
  }
}
/*
async function authenticateUser() {
  const clientSecret = "Inl8Q~jhDg8qQ5jrhBTuQBCQbGdkHmcQLpMqEcTQ";
  const secretID = "ad646418-c15f-44b1-90cc-0af31238d1e6";
  const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  //return noLogin();
  return withLogin();

  async function noLogin() {
    const params = new URLSearchParams();
    params.append("client_id", clientId);
    params.append("client_secret", clientSecret);
    params.append("scope", "https://graph.microsoft.com/.default");
    params.append("grant_type", "client_credentials");
    console.log("params =", params);
    try {
      const response = await fetch(tokenEndpoint, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: params
      });

      const data = await response.json();
      if (data.access_token) {
        console.log("✅ Access Token Retrieved:", data.access_token);
        return data.access_token;
      } else {
        throw new Error("❌ Failed to retrieve access token: " + JSON.stringify(data));
      }
    } catch (error) {
      console.error("Error getting access token:", error);
    }
  }

  function withLogin() {
    const scopes = "Files.ReadWrite";
    const redirectUri = encodeURI("https://login.microsoftonline.com/common/oauth2/nativeclient");
    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}&scope=${scopes}&response_mode=fragment`;
    return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(authUrl, { height: 60, width: 40 }, (result) => {
        console.log("result = ", result);
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
            const token = new URLSearchParams(args.message.split("#")[1]).get("access_token");
            dialog.close();
            debugger;
            if (token) resolve(token);
            else reject("Authentication failed");
          });
        } else {
          reject("Failed to open login dialog.");
        }
      });
    });
  }
}*/

/*
async function getAccessToken() {

  const scopes = ["Files.ReadWrite", "User.Read"];
  const redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient";

  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}&scope=${scopes.join(
    "%20"
  )}&response_mode=fragment`;

  // Open authentication window
  const authWindow = window.open(authUrl, "_blank");

  return new Promise((resolve, reject) => {
    const checkAuth = setInterval(() => {
      try {
        if (authWindow.location.hash) {
          clearInterval(checkAuth);
          const token = new URLSearchParams(authWindow.location.hash.substring(1)).get("access_token");
          authWindow.close();
          resolve(token);
        }
      } catch (e) {}
    }, 1000);
  });
}*/

async function uploadWordDocument(filtered: any[][], fileName: string) {
  //console.log('filtered = ', filtered);
  //const accessToken = await getAccessToken();
  //const accessToken = await authenticateUser();
  const accessToken = await getTokenWithMSAL();
  if (accessToken) {
    console.log("Successfully retrieved token:", accessToken);
    Office.context.ui.messageParent(`Access token: ${accessToken}`);
  } else {
    console.log("Failed to retrieve token.");
  }
  if (!accessToken) return console.log("No access token");

  // Sample Word document content (base64 encoded DOCX)
  const wordContent = "UEsDBBQAAAAIA...";

  const oneDriveUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/Documents/${fileName}.docx:/content`;

  const response = await fetch(oneDriveUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    },
    body: atob(wordContent) // Convert base64 to binary
  });

  if (response.ok) {
    console.log("File uploaded successfully!");
  } else {
    console.error("Upload failed", await response.text());
  }
}

