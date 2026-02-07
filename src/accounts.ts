import { promises as fs } from 'node:fs';
import {
  AccountConfig,
  AccountsConfig,
  getAccountPaths,
  validateAccountId,
} from './config.js';

export interface AccountHealth {
  account: AccountConfig;
  hasCredentialsFile: boolean;
  hasTokenFile: boolean;
  ready: boolean;
}

export function getAccountOrThrow(config: AccountsConfig, accountId: string): AccountConfig {
  validateAccountId(accountId);
  const account = config.accounts.find((item) => item.id === accountId);
  if (!account) {
    throw new Error(`Unknown account "${accountId}".`);
  }
  return account;
}

export function getEnabledAccounts(config: AccountsConfig): AccountConfig[] {
  return config.accounts.filter((account) => account.enabled);
}

export function resolveReadAccounts(config: AccountsConfig, accountId?: string): AccountConfig[] {
  if (accountId && accountId.trim() !== '') {
    const account = getAccountOrThrow(config, accountId);
    if (!account.enabled) {
      throw new Error(`Account "${accountId}" is disabled.`);
    }
    return [account];
  }

  const enabledAccounts = getEnabledAccounts(config);
  if (enabledAccounts.length === 0) {
    throw new Error(
      'No enabled Gmail accounts are configured. Use begin_account_auth and finish_account_auth first.'
    );
  }

  return enabledAccounts;
}

export function resolveWriteAccount(config: AccountsConfig, accountId?: string): AccountConfig {
  if (!accountId || accountId.trim() === '') {
    throw new Error('This tool requires an explicit "account" value.');
  }

  const account = getAccountOrThrow(config, accountId);
  if (!account.enabled) {
    throw new Error(`Account "${accountId}" is disabled.`);
  }

  return account;
}

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

export async function getAccountHealth(
  configRoot: string,
  account: AccountConfig
): Promise<AccountHealth> {
  const paths = getAccountPaths(configRoot, account);
  const [hasCredentialsFile, hasTokenFile] = await Promise.all([
    fileExists(paths.credentialsPath),
    fileExists(paths.tokenPath),
  ]);

  return {
    account,
    hasCredentialsFile,
    hasTokenFile,
    ready: account.enabled && hasCredentialsFile && hasTokenFile,
  };
}
