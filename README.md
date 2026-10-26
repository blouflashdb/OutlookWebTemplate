# OutlookWebTemplate

Template for Outlook Web Add-ins.

| Area | Tooling |
| --- | --- |
| UI | Vue 3, [Nuxt UI](https://ui.nuxt.com) (Vue mode, Tailwind CSS 4) |
| Build | Vite 8 |
| Language | TypeScript 6 with [`@tsconfig/bases`](https://github.com/tsconfig/bases): `strictest`, `node24`, `node-ts` |
| Lint & format | ESLint with [`@antfu/eslint-config`](https://github.com/antfu/eslint-config) (type-aware, erasable syntax only, pnpm rules), run on commit by [lefthook](https://lefthook.dev) |
| Tests | Vitest, Vue Test Utils, jsdom |
| Package manager | pnpm 12 (via Corepack), Node.js 24 |
| CI & updates | Azure DevOps Pipelines, Renovate |
| Office tooling | `office-addin-debugging`, `office-addin-dev-certs`, `office-addin-manifest` |

## Getting started

```bash
corepack enable
```

```bash
pnpm install
```

`pnpm install` also installs the Git hooks (see [Git hooks](#git-hooks)).

Outlook only loads add-ins over HTTPS. Install the trusted localhost certificate once (it asks to add a root certificate to your trust store):

```bash
pnpm certs:install
```

Then check that nothing is missing:

```powershell
.\scripts\check-dev-setup.ps1
```

The script only reads. It checks Node.js (version from `.nvmrc` and `engines`), Corepack, pnpm (the version pinned in `packageManager`), Git, whether the dependencies match `pnpm-lock.yaml`, the lefthook hooks, the trusted localhost certificate and the recommended VS Code extensions (`code` or `code-insiders`). Each problem comes with the command that fixes it, and the script exits with code 1 if a required check fails. It runs in Windows PowerShell 5.1 and PowerShell 7. If your execution policy blocks local scripts, allow them once with `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, or turn on **Settings** > **System** > **For developers** > **PowerShell**.

### Sideloading

How you get the add-in into Outlook depends on your account. Either way it is installed into your mailbox, so it then shows up in every Outlook client you use with that account.

**Any account, including personal outlook.com accounts.** Sideload once by hand, then start the dev server:

1. Open <https://aka.ms/olksideload> (Outlook on the web: **My add-ins** > **Custom add-ins** > **Add a custom add-in** > **Add from file**) and upload `manifest.xml`.
2. Run `pnpm dev`, or start **Outlook on the web (Edge)** from the VS Code debug menu.

**Microsoft 365 work or school account.** `pnpm start` signs in, uploads `manifest.xml` through the Microsoft 365 Agents Toolkit and opens Outlook; `pnpm stop` removes it again. Your tenant must allow custom app uploads. Personal accounts can sign in but the upload service rejects them with `401`.

Open a message and use **Show Task Pane** in the ribbon.

## Scripts

| Script | What it does |
| --- | --- |
| `pnpm dev` | Vite dev server with HMR on `https://localhost:3000` |
| `pnpm start` / `pnpm stop` | Start the dev server and sideload into Outlook / remove the add-in |
| `pnpm build` | Type-check with `vue-tsc`, then build to `dist/` |
| `pnpm preview` | Serve `dist/` on `https://localhost:3000` |
| `pnpm test` | Vitest in watch mode |
| `pnpm test:ci` | Single run with coverage (Cobertura) and a JUnit report |
| `pnpm lint` / `pnpm lint:fix` | ESLint, which also does the formatting |
| `pnpm type-check` | `vue-tsc --build` across all tsconfig projects |
| `pnpm validate` | Validate `manifest.xml` with Microsoft's validation service |

## Project layout

```
manifest.xml            Add-in only (XML) manifest, pointing at https://localhost:3000
taskpane.html           Task pane entry (loads Office.js from the CDN)
commands.html           FunctionFile for ExecuteFunction ribbon buttons
public/assets/          Ribbon and store icons
src/
  taskpane/             Vue app shown in the task pane
  commands/             Ribbon command handlers
  composables/          Office.js wrappers used by the UI
  test-utils/office.ts  Office global stub for unit tests
.azuredevops/           CI and Renovate pipelines
.vscode/                Recommended extensions, settings, Outlook debug config
```

## When you create a project from this template

1. Give `<Id>` in `manifest.xml` a new GUID (`node -e "console.log(crypto.randomUUID())"`), and update `ProviderName`, `DisplayName`, `Description` and `SupportUrl`.
2. Replace the icons in `public/assets/` (16, 32, 64, 80 and 128 px PNGs).
3. Rename the package in `package.json`.

## Building for production

Every URL in `manifest.xml` points at the dev server. `pnpm build` writes `dist/manifest.xml` with `https://localhost:3000/` replaced by `ADDIN_PUBLIC_URL`, the URL `dist/` is hosted at:

```bash
ADDIN_PUBLIC_URL=https://addin.contoso.com/outlook/ pnpm build
```

Assets use relative URLs, so `dist/` can be hosted under any path. Deploy `dist/` to static hosting and give users `dist/manifest.xml` (for example through Integrated Apps in the Microsoft 365 admin center).

## TypeScript

The tsconfigs extend [`@tsconfig/bases`](https://github.com/tsconfig/bases) and are split into projects that `vue-tsc --build` checks together:

- `tsconfig.app.json` covers the add-in code (`strictest` + `node-ts`, DOM libs, Office.js types).
- `tsconfig.vitest.json` covers unit tests (app config plus Node types).
- `tsconfig.node.json` covers `vite.config.ts` and `eslint.config.ts` (`node24` + `node-ts` + `strictest`).

`@tsconfig/node-ts` enables `erasableSyntaxOnly`, so only TypeScript syntax that can be stripped without transformation compiles: no `enum`, `namespace` or constructor parameter properties. ESLint enforces the same rules (`erasableOnly`). The app turns off node-ts's `rewriteRelativeImportExtensions`, because it only matters when `tsc` emits and it rejects `.vue` imports.

TypeScript is held at 6.x. TypeScript 7 (the native port) no longer ships the JavaScript compiler API that `vue-tsc` and `typescript-eslint` depend on. Renovate is configured not to offer it.

## Nuxt UI

Nuxt UI runs in Vue mode through `@nuxt/ui/vite` (see `vite.config.ts`) with `router: false`, because the task pane is a single view. Components and composables such as `UButton` and `useToast` are auto-imported. Their types are generated into `auto-imports.d.ts` and `components.d.ts`, which are committed so type-checking works on a clean checkout. Theme colors are set in the `ui` option. Icons come from the Lucide set (`i-lucide-*`) and are bundled at build time, so Outlook makes no requests to the Iconify API.

## Testing

Office.js only exists inside an Office host, so tests install a minimal `Office` global with `stubOffice()` from `src/test-utils/office.ts`. It stubs only the APIs this add-in uses; extend it as you call more of Office.js.

## Git hooks

[lefthook](https://lefthook.dev) runs the hooks in `lefthook.yml`. Its postinstall installs them into `.git/hooks` on `pnpm install`, and skips that when `CI` is set.

- **pre-commit:** `eslint --fix` on the staged JavaScript, TypeScript, Vue, JSON, YAML and Markdown files. Fixes are staged again automatically (`stage_fixed`), and the commit stops if errors remain. Files ESLint ignores are skipped quietly (`--no-warn-ignored`).

`glob_matcher: doublestar` gives `**` its usual meaning, so `**/*.ts` also matches files in the repository root. To skip the hooks for one commit, use `git commit --no-verify` or set `LEFTHOOK=0`.

## VS Code

Open the folder and install the recommended extensions (Vue, ESLint, Vitest, Azure Pipelines). ESLint formats on save; Prettier is disabled.

The Run and Debug view has two configurations:

- **Outlook on the web (Edge)** starts the dev server and opens Outlook on the web in Edge with the debugger attached, so breakpoints in `src/` work inside the add-in. It works with any account once the add-in is sideloaded. Edge uses a separate profile that remembers your sign-in. The URL is `outlook.live.com` for personal accounts; change it to `outlook.office.com` for work or school accounts.
- **Outlook Desktop (Microsoft 365 account)** runs `pnpm start` and attaches to the add-in's WebView2 in classic Outlook on Windows. It needs a work or school account (see [Sideloading](#sideloading)).

## Azure DevOps

### CI (`.azuredevops/ci.yml`)

Runs on pushes and pull requests to `main`: install (with a cached pnpm store), lint, tests (results and coverage are published to the run), type-check and build. `dist/` is published as the pipeline artifact `addin`.

Set the pipeline variable `ADDIN_PUBLIC_URL` (in the pipeline UI or a variable group) so the built manifest points at your hosting.

### Renovate (`.azuredevops/renovate.yml`)

Azure DevOps has no hosted Renovate app, so a scheduled pipeline runs Renovate (weekdays 03:00 UTC) and opens update PRs. One-time setup:

1. Create a pipeline from `.azuredevops/renovate.yml`.
2. In **Project settings** > **Repositories** > this repo > **Security**, give `<project> Build Service (<org>)` **Contribute**, **Contribute to pull requests**, **Create branch** and **Create tag**.
3. Add a secret pipeline variable `RENOVATE_GITHUB_COM_TOKEN`: a GitHub token with no scopes, used for changelogs and to avoid GitHub rate limits.

`renovate.json` extends `config:recommended`, bumps ranges in `package.json`, maintains the lockfile weekly, updates Azure Pipelines tasks, and keeps the Renovate version pinned in the pipeline up to date.

## pnpm supply-chain settings

`pnpm-workspace.yaml` holds the settings `@antfu/eslint-config` enforces:

- `trustPolicy: no-downgrade` rejects a package version published with weaker provenance than an earlier one, a sign of a possible takeover. `trustPolicyExclude` lists the known exceptions: Microsoft's sideload toolchain and a few old maintenance releases.
- `allowBuilds` makes dependency install scripts opt-in. Only `keytar` (used by the Microsoft 365 sign-in in `pnpm start`) and `lefthook` (installs the Git hooks) run theirs.

`lefthook` is pinned to 2.2.1. Versions 2.1.16 to 2.2.1 were published from GitHub Actions with provenance, but 2.2.2 was uploaded by hand without it, so the trust policy rejects it. Renovate's update to 2.2.2 fails for the same reason; take the next release that has provenance again.

`publicHoistPattern` hoists `@microsoft/m365agentstoolkit-cli` to the root of `node_modules`. It is a peer dependency of `office-addin-debugging`, which runs its `atk` binary to sideload. pnpm installs the peer automatically but only links binaries of root-level packages into `node_modules/.bin`; without the hoist, `pnpm start` fails with `'atk' is not recognized`.
