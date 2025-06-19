import { AccessToken, TokenCredential, ClientSecretCredential, InteractiveBrowserCredential, DeviceCodeCredential, DeviceCodeInfo } from "@azure/identity";
import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import { logger } from "./logger.js";

// Simple authentication provider that works with Azure Identity TokenCredential
export class TokenCredentialAuthProvider implements AuthenticationProvider {
  private credential: TokenCredential;

  constructor(credential: TokenCredential) {
    this.credential = credential;
  }

  async getAccessToken(): Promise<string> {
    const token = await this.credential.getToken("https://graph.microsoft.com/.default");
    if (!token) {
      throw new Error("Failed to acquire access token");
    }
    return token.token;
  }
}

export interface TokenBasedCredential extends TokenCredential {
  getToken(scopes: string | string[]): Promise<AccessToken | null>;
}

export class ClientProvidedTokenCredential implements TokenBasedCredential {
  private accessToken: string;
  private expiresOn: Date;

  constructor(accessToken: string, expiresOn?: Date) {
    this.accessToken = accessToken;
    this.expiresOn = expiresOn || new Date(Date.now() + 3600000); // Default 1 hour
  }

  async getToken(scopes: string | string[]): Promise<AccessToken | null> {
    if (this.expiresOn <= new Date()) {
      logger.error("Access token has expired");
      return null;
    }

    return {
      token: this.accessToken,
      expiresOnTimestamp: this.expiresOn.getTime()
    };
  }

  updateToken(accessToken: string, expiresOn?: Date): void {
    this.accessToken = accessToken;
    this.expiresOn = expiresOn || new Date(Date.now() + 3600000);
    logger.info("Access token updated successfully");
  }

  isExpired(): boolean {
    return this.expiresOn <= new Date();
  }

  getExpirationTime(): Date {
    return this.expiresOn;
  }
}

export enum AuthMode {
  ClientCredentials = "client_credentials",
  ClientProvidedToken = "client_provided_token", 
  Interactive = "interactive"
}

export interface AuthConfig {
  mode: AuthMode;
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
  accessToken?: string;
  expiresOn?: Date;
  redirectUri?: string;
}

export class AuthManager {
  private credential: TokenCredential | null = null;
  private config: AuthConfig;

  constructor(config: AuthConfig) {
    this.config = config;
  }

  async initialize(): Promise<void> {
    switch (this.config.mode) {
      case AuthMode.ClientCredentials:
        if (!this.config.tenantId || !this.config.clientId || !this.config.clientSecret) {
          throw new Error("Client credentials mode requires tenantId, clientId, and clientSecret");
        }
        logger.info("Initializing Client Credentials authentication");
        this.credential = new ClientSecretCredential(
          this.config.tenantId,
          this.config.clientId,
          this.config.clientSecret
        );
        break;

      case AuthMode.ClientProvidedToken:
        if (!this.config.accessToken) {
          throw new Error("Client provided token mode requires accessToken");
        }
        logger.info("Initializing Client Provided Token authentication");
        this.credential = new ClientProvidedTokenCredential(
          this.config.accessToken,
          this.config.expiresOn
        );
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
        } catch (error) {
          // Fallback to Device Code flow
          logger.info("Interactive browser failed, falling back to device code flow");
          this.credential = new DeviceCodeCredential({
            tenantId: this.config.tenantId,
            clientId: this.config.clientId,
            userPromptCallback: (info: DeviceCodeInfo) => {
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

  updateAccessToken(accessToken: string, expiresOn?: Date): void {
    if (this.config.mode === AuthMode.ClientProvidedToken && this.credential instanceof ClientProvidedTokenCredential) {
      this.credential.updateToken(accessToken, expiresOn);
    } else {
      throw new Error("Token update only supported in client provided token mode");
    }
  }

  private async testCredential(): Promise<void> {
    if (!this.credential) {
      throw new Error("Credential not initialized");
    }

    try {
      const token = await this.credential.getToken("https://graph.microsoft.com/.default");
      if (!token) {
        throw new Error("Failed to acquire token");
      }
      logger.info("Authentication successful");
    } catch (error) {
      logger.error("Authentication test failed", error);
      throw error;
    }
  }
  getGraphAuthProvider(): TokenCredentialAuthProvider {
    if (!this.credential) {
      throw new Error("Authentication not initialized");
    }

    return new TokenCredentialAuthProvider(this.credential);
  }

  getAzureCredential(): TokenCredential {
    if (!this.credential) {
      throw new Error("Authentication not initialized");
    }
    return this.credential;
  }

  getAuthMode(): AuthMode {
    return this.config.mode;
  }

  isClientCredentials(): boolean {
    return this.config.mode === AuthMode.ClientCredentials;
  }

  isClientProvidedToken(): boolean {
    return this.config.mode === AuthMode.ClientProvidedToken;
  }

  isInteractive(): boolean {
    return this.config.mode === AuthMode.Interactive;
  }

  getTokenStatus(): { isExpired: boolean; expiresOn?: Date } {
    if (this.credential instanceof ClientProvidedTokenCredential) {
      return {
        isExpired: this.credential.isExpired(),
        expiresOn: this.credential.getExpirationTime()
      };
    }
    return { isExpired: false };
  }
}
