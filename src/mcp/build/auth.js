import { ClientSecretCredential, InteractiveBrowserCredential, DeviceCodeCredential } from "@azure/identity";
import { logger } from "./logger.js";
// Simple authentication provider that works with Azure Identity TokenCredential
export class TokenCredentialAuthProvider {
    credential;
    constructor(credential) {
        this.credential = credential;
    }
    async getAccessToken() {
        const token = await this.credential.getToken("https://graph.microsoft.com/.default");
        if (!token) {
            throw new Error("Failed to acquire access token");
        }
        return token.token;
    }
}
export class ClientProvidedTokenCredential {
    accessToken;
    expiresOn;
    constructor(accessToken, expiresOn) {
        this.accessToken = accessToken;
        this.expiresOn = expiresOn || new Date(Date.now() + 3600000); // Default 1 hour
    }
    async getToken(scopes) {
        if (this.expiresOn <= new Date()) {
            logger.error("Access token has expired");
            return null;
        }
        return {
            token: this.accessToken,
            expiresOnTimestamp: this.expiresOn.getTime()
        };
    }
    updateToken(accessToken, expiresOn) {
        this.accessToken = accessToken;
        this.expiresOn = expiresOn || new Date(Date.now() + 3600000);
        logger.info("Access token updated successfully");
    }
    isExpired() {
        return this.expiresOn <= new Date();
    }
    getExpirationTime() {
        return this.expiresOn;
    }
}
export var AuthMode;
(function (AuthMode) {
    AuthMode["ClientCredentials"] = "client_credentials";
    AuthMode["ClientProvidedToken"] = "client_provided_token";
    AuthMode["Interactive"] = "interactive";
})(AuthMode || (AuthMode = {}));
export class AuthManager {
    credential = null;
    config;
    constructor(config) {
        this.config = config;
    }
    async initialize() {
        switch (this.config.mode) {
            case AuthMode.ClientCredentials:
                if (!this.config.tenantId || !this.config.clientId || !this.config.clientSecret) {
                    throw new Error("Client credentials mode requires tenantId, clientId, and clientSecret");
                }
                logger.info("Initializing Client Credentials authentication");
                this.credential = new ClientSecretCredential(this.config.tenantId, this.config.clientId, this.config.clientSecret);
                break;
            case AuthMode.ClientProvidedToken:
                if (!this.config.accessToken) {
                    throw new Error("Client provided token mode requires accessToken");
                }
                logger.info("Initializing Client Provided Token authentication");
                this.credential = new ClientProvidedTokenCredential(this.config.accessToken, this.config.expiresOn);
                break;
            case AuthMode.Interactive:
                if (!this.config.tenantId || !this.config.clientId) {
                    throw new Error("Interactive mode requires tenantId and clientId");
                }
                logger.info("Initializing Interactive authentication");
                try {
                    // Try Interactive Browser first
                    this.credential = new InteractiveBrowserCredential({
                        tenantId: this.config.tenantId,
                        clientId: this.config.clientId,
                        redirectUri: this.config.redirectUri || "http://localhost:3000",
                    });
                }
                catch (error) {
                    // Fallback to Device Code flow
                    logger.info("Interactive browser failed, falling back to device code flow");
                    this.credential = new DeviceCodeCredential({
                        tenantId: this.config.tenantId,
                        clientId: this.config.clientId,
                        userPromptCallback: (info) => {
                            console.log(`\n🔐 Authentication Required:`);
                            console.log(`Please visit: ${info.verificationUri}`);
                            console.log(`And enter code: ${info.userCode}\n`);
                            return Promise.resolve();
                        },
                    });
                }
                break;
            default:
                throw new Error(`Unsupported authentication mode: ${this.config.mode}`);
        }
        // Test the credential
        await this.testCredential();
    }
    updateAccessToken(accessToken, expiresOn) {
        if (this.config.mode === AuthMode.ClientProvidedToken && this.credential instanceof ClientProvidedTokenCredential) {
            this.credential.updateToken(accessToken, expiresOn);
        }
        else {
            throw new Error("Token update only supported in client provided token mode");
        }
    }
    async testCredential() {
        if (!this.credential) {
            throw new Error("Credential not initialized");
        }
        try {
            const token = await this.credential.getToken("https://graph.microsoft.com/.default");
            if (!token) {
                throw new Error("Failed to acquire token");
            }
            logger.info("Authentication successful");
        }
        catch (error) {
            logger.error("Authentication test failed", error);
            throw error;
        }
    }
    getGraphAuthProvider() {
        if (!this.credential) {
            throw new Error("Authentication not initialized");
        }
        return new TokenCredentialAuthProvider(this.credential);
    }
    getAzureCredential() {
        if (!this.credential) {
            throw new Error("Authentication not initialized");
        }
        return this.credential;
    }
    getAuthMode() {
        return this.config.mode;
    }
    isClientCredentials() {
        return this.config.mode === AuthMode.ClientCredentials;
    }
    isClientProvidedToken() {
        return this.config.mode === AuthMode.ClientProvidedToken;
    }
    isInteractive() {
        return this.config.mode === AuthMode.Interactive;
    }
    getTokenStatus() {
        if (this.credential instanceof ClientProvidedTokenCredential) {
            return {
                isExpired: this.credential.isExpired(),
                expiresOn: this.credential.getExpirationTime()
            };
        }
        return { isExpired: false };
    }
}
