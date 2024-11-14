//MSAL configuration
const msalConfig = {
    auth: {
        clientId: '<clienteID-Register-Aplication>',
        // comment out if you use a multi-tenant AAD app
        authority: 'https://login.microsoftonline.com/<ID-Locatario-Register-Aplication>',
        redirectUri: 'http://localhost:8080'
    }
};
const msalRequest = { scopes: [] };
function ensureScope (scope) {
    if (!msalRequest.scopes.some((s) => s.toLowerCase() === scope.toLowerCase())) {
        msalRequest.scopes.push(scope);
    }
}
//Initialize MSAL client
const msalClient = new msal.PublicClientApplication(msalConfig);

// Log the user in
async function signIn() {
    const authResult = await msalClient.loginPopup(msalRequest);
    sessionStorage.setItem('msalAccount', authResult.account.username);
}
//Get token from Graph
async function getToken() {
    let account = sessionStorage.getItem('msalAccount');
    if (!account) {
        throw new Error(
            'User info cleared from session. Please sign out and sign in again.');
    }
    try {
        // First, attempt to get the token silently
        const silentRequest = {
            scopes: msalRequest.scopes,
            account: msalClient.getAccountByUsername(account)
        };

        const silentResult = await msalClient.acquireTokenSilent(silentRequest);
        return silentResult.accessToken;
    } catch (silentError) {
        // If silent requests fails with InteractionRequiredAuthError,
        // attempt to get the token interactively
        if (silentError instanceof msal.InteractionRequiredAuthError) {
            const interactiveResult = await msalClient.acquireTokenPopup(msalRequest);
            return interactiveResult.accessToken;
        } else {
            throw silentError;
        }
    }
}
async function getEmails(){
    ensureScope('mail.read');
    
    return await graphClient
    .api('/me/messages')    
    .select('subject,receivedDateTime')
    .orderby('receivedDateTime desc')
    .top(10)
    .get(); 
}
async function displayEmail() {
    var emails = await getEmails();
    if (!emails || emails.value.length < 1) {
      return;
    }
  
    document.getElementById('displayEmail').style = 'display: none';
  
    var emailsUl = document.getElementById('emails');
    emails.value.forEach(email => {
      var emailLi = document.createElement('li');
      emailLi.innerText = `${email.subject} (${new Date(email.receivedDateTime).toLocaleString()})`;
      emailsUl.appendChild(emailLi);
    });
  }
