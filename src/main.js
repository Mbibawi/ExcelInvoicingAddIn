"use strict";
Office.onReady((info) => {
    // Office is ready --
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
}
;
function selectForm(id) {
    showForm(id);
}
async function showForm(id) {
    const formHtml = `<div id="filterForm" >
    <label for="client" > Client: </label>
    <input type="text" id="_client" name="client"><br><br>
    <input list="clients" id="client" name="client" data-index = "0" autocomplete="on"><br><br>
    <label for="matter" > Affaire: </label>
    <input type ="text" id ="_matter" name="affaire"><br><br>
    <input list="matters" id="matter" name="matter" data-index = "1" autocomplete="on"><br><br>
    <label for="nature" > Nature: </label>
    <input type ="text" id ="_nature" name="nature"><br><br>
    <input list="natures" id="nature" name="nature" data-index="2" autocomplete="on"><br><br>
    <label for="date" > Date: </label>
    <input type="date" id="date" name="date" data-index="3"><br><br>
  </div>`;
    const form = document.getElementById("form");
    form.innerHTML = '';
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const table = sheet.tables.getItem('LivreJournal');
        const title = table.getHeaderRowRange();
        title.load('values');
        await context.sync();
        //form.innerHTML = formHtml;
        insertInputsAndLables(['client', 'matter', 'nature', 'date'], title.values[0]);
        if (id === 'invoice')
            form.innerHTML += `<button onclick="filter()"> Filter Table</button>`;
        if (id !== 'entry')
            return;
        const otherHtml = `
  <label for="adress" > Adresse: </label>
  <input type ="text" id ="_adress" name="adress"><br><br>
  <input list="adress" id="adress" name="adress" data-index="10" autocomplete="on"><br><br>
  <label for="adress" > Moyen de paiement: </label>
  <input type ="text" id ="_payment" name="payment"><br><br>
  <input list="payment" id="payment" name="payment" data-index="5" autocomplete="on"><br><br>
  <label for="amount" > Montant: </label>
  <input type ="text" id ="_amout" name="amount"><br><br>
  <input list="amount" id="amount" name="amount" data-index="8" autocomplete="on"><br><br>
  <label for="vat" > TVA: </label>
  <input type ="text" id ="_vat" name="vat"><br><br>
  <input list="vat" id="vat" name="vat" data-index=autocomplete="on"><br><br>
  <label for="account" > Bank account: </label>
  <input type ="text" id ="_account" name="account"><br><br>
  <input list="account" id="account" name="account" data-index="6" autocomplete="on"><br><br>
  <label for="payee" > Third Party: </label>
  <input type ="text" id ="_payee" name="payee"><br><br>
  <input list="payee" id="payee" name="payee" data-index="7" autocomplete="on"><br><br>
  <label for="description" > Description: </label>
  <input type ="text" id ="_description" name="description"><br><br>
  <input list="description" id="description" name="description" data-index="8" autocomplete="off"><br><br>
  <button onclick="addEntry()"> Ajouter </button>
  `;
        form.innerHTML += otherHtml;
        // if(id === 'insert') getInputs([null, null, null, null, 'amount', 'vat', 'account', 'payment', 'payee', 'description', 'adress' ], title.values[0])
    });
    function insertInputsAndLables(ids, title) {
        ids.forEach(id => {
            if (!id)
                return;
            const input = document.createElement('input');
            const index = ids.indexOf(id);
            input.type = "text";
            if (id === 'date')
                input.type = 'date';
            input.setAttribute('list', id + 's'),
                input.id = id;
            input.name = id;
            input.dataset.index = index.toString();
            input.autocomplete = "on";
            const label = document.createElement('label');
            label.htmlFor = id;
            label.innerText = title[index];
            form.appendChild(label);
            form.appendChild(input);
            form.appendChild(document.createElement('br'));
            form.appendChild(document.createElement('br'));
        });
        addOnChange();
    }
    function addOnChange(inputs) {
        if (!inputs || inputs.length < 1)
            inputs = Array.from(document.getElementsByTagName('input'));
        console.log('inputs = ', inputs);
        inputs.forEach(input => input.addEventListener('change', async () => await inputOnChange(input)));
    }
    async function inputOnChange(input) {
        debugger;
        const visible = await filterTable(undefined, [[Number(input.dataset.index), getArray(input.value)]]);
        if (!visible)
            return;
        const nextInput = input.nextElementSibling;
        if (!nextInput)
            return;
        const nextColumnValues = new Set(visible.map(row => row[Number(nextInput.dataset.index)]));
        console.log('unique values = ', nextColumnValues);
        // Create a new datalist element
        let dataList = Array.from(document.getElementsByTagName('datalist')).find(list => list.id === input.id + "s");
        if (dataList)
            dataList.remove();
        dataList = document.createElement('datalist');
        //@ts-ignore
        dataList.id = input.id + 's';
        // Append options to the datalist
        Array.from(nextColumnValues).forEach(option => {
            const optionElement = document.createElement('option');
            optionElement.value = option;
            dataList?.appendChild(optionElement);
        });
        // Attach the datalist to the body or a specific element
        document.body.appendChild(dataList);
    }
}
/**
 * Filters the Excel table based on a criteria
 * @param {[[number, string[]]]} criteria - the first element is the column index, the second element is the values[] based on which the column will be filtered
 */
async function filterTable(tableName = 'LivreJournal', criteria, clearFilter = false) {
    return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const table = sheet.tables.getItem(tableName);
        if (clearFilter)
            table.autoFilter.clearCriteria();
        criteria.forEach(column => filterColumn(column[0], column[1]));
        function filterColumn(index, filter) {
            table.columns.getItemAt(index).filter.applyValuesFilter(filter);
        }
        await context.sync();
        const range = table.getDataBodyRange().getVisibleView();
        range.load('values');
        return range.values;
        //await createWordDocument(range.values);
        //@ts-ignore
        uploadWordDocument(range.values, criteria.client.join(",") + "_Invoice" + "THEDATE");
    });
}
/**
 * Converts the ',' separated text in the input into an array
 * @param value
 * @returns {string[]}
 */
function getArray(value) {
    const array = value.replaceAll(', ', ',')
        .replaceAll(' ,', ',')
        .split(',');
    return array.filter((el) => el);
}
async function filter() {
    const client = document.getElementById("client");
    const matter = document.getElementById("affaire");
    const nature = document.getElementById("nature");
    const date = document.getElementById("date");
    const criteria = [
        [0, getArray(client.value) || undefined],
        [1, getArray(matter.value) || undefined],
        [2, getArray(nature.value) || undefined],
        [3, getArray(date.value) || undefined]
    ];
    filterTable(undefined, criteria, true);
}
async function addEntry(tableName = 'LivreJournal') {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const table = sheet.tables.getItem(tableName);
        //const columns = ['client', 'matter', 'nature', 'date', 'year', 'startTime', 'endTime', 'totalTime', 'hourlyRate', 'amount', 'vat', 'payment', 'account', 'payee', 'description', 'link', 'adress']
        const inputs = Array.from(document.getElementsByTagName('input')).filter(input => input.dataset.index);
        const titleRow = table.rows.getItemAt(0).values;
        const row = titleRow[0].map(cell => {
            const input = inputs.find(el => Number(el.dataset.index) === titleRow.indexOf(cell));
            if (!input)
                return '';
            return input.value;
        });
        /*
        columns.map(column => {
          
        });
        columns.length = Number(table.columns.getCount());
        columns.forEach
        const row = columns.map(id => {
          const input = document.getElementById(id) as HTMLInputElement;
          if (!input) return '';
          return input.value
        })*/
        table.rows.add(-1, [row], true);
        await context.sync();
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
    async function checkExistingAuthContext() {
        try {
            const account = msalInstance.getAllAccounts()[0];
            if (account) {
                return acquireTokenSilently(account);
            }
            else {
                //return loginAndGetToken();
                //openLoginWindow()
                return getOfficeToken();
                //return getTokenWithSSO('minabibawi@gmail.com')
                //return credentitalsToken()
            }
        }
        catch (error) {
            console.error("Error checking auth context:", error);
            return undefined;
        }
    }
    async function credentitalsToken() {
        const msalConfig = {
            auth: {
                clientId: clientId,
                authority: `https://login.microsoftonline.com/${tenantId}`,
                //clientSecret: clientSecret,
            }
        };
        //@ts-ignore
        const cca = new msal.application.ConfidentialClientApplication(msalConfig);
        const tokenRequest = {
            scopes: ["Files.ReadWrite"],
        };
        try {
            const response = await cca.acquireTokenByClientCredential(tokenRequest);
            return response.accessToken;
        }
        catch (error) {
            console.log('Error acquiring Token: ', error);
            return null;
        }
    }
    async function getOfficeToken() {
        try {
            //@ts-ignore
            return await OfficeRuntime.auth.getAccessToken();
        }
        catch (error) {
            console.log("Error : ", error);
        }
    }
    async function getTokenWithSSO(email) {
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
        }
        catch (error) {
            console.error("SSO silent authentication failed:", error);
            return null;
        }
    }
    function openLoginWindow() {
        const loginUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}&scope=https://graph.microsoft.com/.default`;
        // Open in a new window (only works if triggered by user action)
        const authWindow = window.open(loginUrl, "_blank", "width=500,height=600");
        if (!authWindow) {
            console.error("Popup blocked! Please allow popups.");
        }
    }
    // Function to handle login and acquire token
    async function loginAndGetToken() {
        try {
            const loginRequest = {
                scopes: ["Files.ReadWrite"] // OneDrive scopes
            };
            await msalInstance.loginRedirect(loginRequest);
            return handleRedirectResponse();
        }
        catch (error) {
            console.error("Login error:", error);
            return undefined;
        }
        // Function to handle redirect response
        async function handleRedirectResponse() {
            try {
                const authResult = await msalInstance.handleRedirectPromise();
                if (authResult && authResult.accessToken) {
                    console.log("Access token:", authResult.accessToken);
                    return authResult.accessToken;
                }
            }
            catch (error) {
                console.error("Redirect handling error:", error);
            }
            return undefined;
        }
    }
    // Function to get access token silently
    async function acquireTokenSilently(account) {
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
        }
        catch (error) {
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
async function uploadWordDocument(filtered, fileName) {
    //console.log('filtered = ', filtered);
    //const accessToken = await getAccessToken();
    //const accessToken = await authenticateUser();
    const accessToken = await getTokenWithMSAL();
    if (accessToken) {
        console.log("Successfully retrieved token:", accessToken);
        Office.context.ui.messageParent(`Access token: ${accessToken}`);
    }
    else {
        console.log("Failed to retrieve token.");
    }
    if (!accessToken)
        return console.log("No access token");
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
    }
    else {
        console.error("Upload failed", await response.text());
    }
}
//# sourceMappingURL=main.js.map