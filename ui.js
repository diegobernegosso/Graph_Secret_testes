  
async function displayUI() {  

    const settings = {
        clientId: '063587f2-3fe2-49f0-aeaf-c8bbb07336ec',
        clientSecret: 'xxxxxxxxxxxxxxxxxxxxxxxxx',
        tenantId: '09800116-7c8b-4d0a-9479-a83ab1e9c7fa',
    };
    
    initializeGraphForAppOnlyAuth(settings);

    async function displayAccessTokenAsync() {
        try {
            const appOnlyToken = await getAppOnlyTokenAsync();
            console.log(`App-only token: ${appOnlyToken}`);
        } catch (err) {
            console.log(`Error getting app-only access token: ${err}`);
        }
    }











    await signIn();

    // Display info from user profile
    const user = await getUser();



    var userName = document.getElementById('userName');
    userName.innerText = user.displayName;  

    // Hide login button and initial UI
    var signInButton = document.getElementById('signin');
    signInButton.style = "display: none";
    var content = document.getElementById('content');
    content.style = "display: block";
}
let _settings = undefined;
let _clientSecretCredential = undefined;
let _appClient = undefined;

function initializeGraphForAppOnlyAuth(settings) {
    // Ensure settings isn't null
    if (!settings) {
        throw new Error('Settings cannot be undefined');
    }

    _settings = settings;

    // Ensure settings isn't null
    if (!_settings) {
        throw new Error('Settings cannot be undefined');
    }

    if (!_clientSecretCredential) {
        _clientSecretCredential = new indentity.ClientSecretCredential(
            _settings.tenantId,
            _settings.clientId,
            _settings.clientSecret,
        );
    }

    if (!_appClient) {
        const authProvider = new TokenCredentialAuthenticationProvider(
            _clientSecretCredential,
            {
                scopes: ['https://graph.microsoft.com/.default'],
            },
        );

        _appClient = Client.initWithMiddleware({
            authProvider: authProvider,
        });
    }
}

async function getAppOnlyTokenAsync() {
    // Ensure credential isn't undefined
    if (!_clientSecretCredential) {
        throw new Error('Graph has not been initialized for app-only auth');
    }

    // Request token with given scopes
    const response = await _clientSecretCredential.getToken([
        'https://graph.microsoft.com/.default',
    ]);
    return response.token;
}
