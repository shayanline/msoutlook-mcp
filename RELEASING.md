# Releasing

`main` is the source of truth. Releases are published to npm as
[`msoutlook-mcp`](https://www.npmjs.com/package/msoutlook-mcp) by the
`publish.yml` workflow, which runs on a **published GitHub Release** using npm
Trusted Publishing (OIDC, no token).

## Steps

1. Make sure `main` is green (typecheck, tests, build) and up to date.
2. Bump the version and update the changelog:
   - Set the new version in `package.json` (follow semver).
   - Move the `Unreleased` notes in `CHANGELOG.md` under a new `## [x.y.z] - YYYY-MM-DD` heading and update the compare links.
   - Commit, e.g. `chore: release vX.Y.Z`, and push to `main`.
3. Create the release on GitHub:
   ```bash
   gh release create vX.Y.Z -R shayanline/msoutlook-mcp --target main \
     --title "vX.Y.Z" --notes "...release notes..."
   ```
   Publishing the release triggers `publish.yml`, which builds, type checks, tests, and publishes to npm.
4. Confirm: `npm view msoutlook-mcp version` shows the new version.

The tag (`vX.Y.Z`) is created by the GitHub Release, so there is no separate
`git tag` step.
