import { promises as fs } from 'node:fs';
import { ensureConfigLayout, getAccountPaths, getConfigRoot, getDefaultAccountPaths, loadAccountsConfig, saveAccountsConfig, upsertAccount, validateAccountId, } from './config.js';
import { getAccountHealth, getAccountOrThrow } from './accounts.js';
import { GmailAccountClient, exchangeCodeForToken, generateAuthUrlFromCredentials, readCredentialsFile, } from './gmail-client.js';
/** Shared account selection, client creation, and OAuth state for MCP and CLI. */
export class AccountSession {
    configRoot;
    constructor(configRoot = getConfigRoot()) {
        this.configRoot = configRoot;
    }
    async loadConfig() {
        return loadAccountsConfig(this.configRoot);
    }
    async getClientForAccount(account) {
        return GmailAccountClient.create(this.configRoot, account);
    }
    /** Reports file availability only; never reads or returns credential/token contents. */
    async listAccounts() {
        const config = await this.loadConfig();
        const accounts = await Promise.all(config.accounts.map(async (account) => {
            const health = await getAccountHealth(this.configRoot, account);
            return {
                id: account.id,
                email: account.email,
                ...(account.displayName === undefined ? {} : { displayName: account.displayName }),
                enabled: account.enabled,
                hasCredentialsFile: health.hasCredentialsFile,
                hasTokenFile: health.hasTokenFile,
                ready: health.ready,
                status: health.ready ? 'ready' : account.enabled ? 'needs-auth-files' : 'disabled',
            };
        }));
        return { defaultAccount: config.defaultAccount, accounts };
    }
    async parseCredentialsInput(args, credentialsPath) {
        if (args.credentials_json !== undefined) {
            if (typeof args.credentials_json === 'string') {
                return JSON.parse(args.credentials_json);
            }
            if (typeof args.credentials_json === 'object' && args.credentials_json !== null) {
                return args.credentials_json;
            }
            throw new Error('credentials_json must be either a JSON string or object.');
        }
        if (args.credentials_path && args.credentials_path.trim() !== '') {
            const raw = await fs.readFile(args.credentials_path, 'utf8');
            return JSON.parse(raw);
        }
        return readCredentialsFile(credentialsPath);
    }
    async beginAuth(args) {
        if (!args.account_id)
            throw new Error('account_id is required.');
        if (!args.email)
            throw new Error('email is required.');
        validateAccountId(args.account_id);
        await ensureConfigLayout(this.configRoot);
        const paths = getDefaultAccountPaths(this.configRoot, args.account_id);
        await fs.mkdir(paths.accountDir, { recursive: true });
        const credentials = await this.parseCredentialsInput(args, paths.credentialsPath);
        const { authUrl } = generateAuthUrlFromCredentials(credentials);
        await fs.writeFile(paths.credentialsPath, `${JSON.stringify(credentials, null, 2)}\n`, 'utf8');
        const config = upsertAccount(await this.loadConfig(), {
            id: args.account_id,
            email: args.email,
            displayName: args.display_name,
            enabled: false,
            credentialPath: paths.credentialsPath,
            tokenPath: paths.tokenPath,
        });
        await saveAccountsConfig(this.configRoot, config);
        return { accountId: args.account_id, email: args.email, authUrl };
    }
    async finishAuth(args) {
        if (!args.account_id)
            throw new Error('account_id is required.');
        if (!args.authorization_code)
            throw new Error('authorization_code is required.');
        validateAccountId(args.account_id);
        let config = await this.loadConfig();
        const account = getAccountOrThrow(config, args.account_id);
        const paths = getAccountPaths(this.configRoot, account);
        const credentials = await readCredentialsFile(paths.credentialsPath);
        const tokens = await exchangeCodeForToken(credentials, args.authorization_code);
        if (!tokens.access_token && !tokens.refresh_token) {
            throw new Error('OAuth exchange succeeded but no token payload was returned.');
        }
        await fs.writeFile(paths.tokenPath, `${JSON.stringify(tokens, null, 2)}\n`, 'utf8');
        const updatedAccount = {
            ...account,
            enabled: true,
            credentialPath: paths.credentialsPath,
            tokenPath: paths.tokenPath,
        };
        const tempConfig = upsertAccount(config, updatedAccount);
        await saveAccountsConfig(this.configRoot, tempConfig);
        const refreshedAccount = getAccountOrThrow(tempConfig, args.account_id);
        const client = await this.getClientForAccount(refreshedAccount);
        let profileEmail = refreshedAccount.email;
        try {
            profileEmail = await client.getProfileEmail();
        }
        catch {
            profileEmail = refreshedAccount.email;
        }
        config = upsertAccount(tempConfig, {
            ...refreshedAccount,
            email: profileEmail,
            enabled: true,
        });
        if (!config.defaultAccount)
            config.defaultAccount = args.account_id;
        await saveAccountsConfig(this.configRoot, config);
        return { accountId: args.account_id, email: profileEmail, enabled: true };
    }
}
//# sourceMappingURL=account-session.js.map