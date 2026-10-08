#!/usr/bin/env node
// No arguments keeps existing MCP integrations and hosted deployments unchanged.
const args = process.argv.slice(2);
if (args.length === 0 || (args.length === 1 && args[0] === 'mcp')) {
    const { runMcp } = await import('./mcp.js');
    runMcp().catch((error) => {
        console.error('[ghub] Fatal error:', error instanceof Error ? error.message : 'Unknown failure');
        process.exitCode = 1;
    });
}
else {
    const { runCli } = await import('./cli.js');
    process.exitCode = await runCli(args);
}
export {};
//# sourceMappingURL=index.js.map