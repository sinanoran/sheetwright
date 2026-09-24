# Releasing

A release is a tag. `.github/workflows/release.yml` reads the version from the
tag name, so there is no version in a file to forget to bump and no way for the
tag and the published package to disagree.

```bash
git tag v0.1.0
git push origin v0.1.0
```

The workflow builds, runs the tests on all three target frameworks, packs,
publishes to nuget.org and creates the GitHub release with the `.nupkg` and
`.snupkg` attached.

## Publishing uses Trusted Publishing, not an API key

nuget.org issues a short-lived key in exchange for a GitHub OIDC token, so this
repository stores no publishing credential at all. There is nothing to rotate
and nothing to leak.

The setup is one policy on nuget.org and one repository secret here.

### The policy, on nuget.org

Sign in, then **your username → Trusted Publishing → add a policy**:

| Field | Value |
|---|---|
| Policy Name | `sheetwright-release` — a label for your own list; nothing matches against it |
| Repository Owner | `sinanoran` |
| Repository | `sheetwright` |
| Workflow File | `release.yml` — the file name only, not the path |
| Environment | leave empty; this workflow uses no GitHub environment |

Under **Scopes**, allow publishing **new packages** as well as new versions, with
the glob `Sheetwright*`. Version 0.1.0 is the first push of an id that does not
exist on nuget.org yet, and a policy scoped to existing packages only will
refuse it.

A policy on a public repository is active immediately. (On a private one it
starts as *pending full activation* for seven days, because nuget.org needs the
repository and owner ids from a real publish to pin the policy against somebody
deleting the repository and recreating it under the same name.)

### The secret, here

`NuGet/login@v1` needs the nuget.org **profile name** — not the email address —
to know whose policy to match:

```bash
gh secret set NUGET_USER --repo sinanoran/sheetwright
```

It is not really a secret, but nuget.org's own documentation recommends passing
it as one, and there is no cost to doing so.

## Before tagging

- `CHANGELOG.md` has a section for the version, and the link definitions at the
  bottom point at it.
- CI is green on `main` for the commit being tagged. The release workflow runs
  the tests again, but finding out at push time is worse than finding out before.

## If the push fails

- **`--skip-duplicate` swallowed it.** The push step passes this flag, so
  re-tagging an already-published version succeeds without republishing. Check
  nuget.org rather than the exit code.
- **The key expired.** It is valid for one hour and is requested in the step
  immediately before the push, so this means something in front of it took
  longer than an hour. Re-run the job.
- **The policy did not match.** The workflow file name, repository owner and
  repository in the policy are compared literally. Renaming this file means
  editing the policy.
