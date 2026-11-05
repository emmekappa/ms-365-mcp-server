# ms-365-mcp-server Helm chart

Runs the server in HTTP mode with OAuth, as a remote MCP connector for many users.
Each user signs in with their own Microsoft account; the server calls Graph with
that user's delegated token.

```sh
helm install ms365-mcp oci://ghcr.io/emmekappa/charts/ms-365-mcp-server \
  --set existingSecret=ms365-mcp \
  --set config.publicUrl=https://mcp.example.com \
  --set config.allowedRedirectUris[0]=https://claude.ai/api/mcp/auth_callback
```

The Secret named by `existingSecret` holds `MS365_MCP_CLIENT_ID`,
`MS365_MCP_CLIENT_SECRET` and `MS365_MCP_TENANT_ID` of the Entra app. Every
setting is in [values.yaml](values.yaml).

The chart is published by `.github/workflows/publish-fork.yml` when the
`version` in `Chart.yaml` changes. Bump it with every change to the chart.
