# Security Policy

## Supported versions

| Version | Branch | Status |
|---|---|---|
| 4.0 beta | `v4` | Pre-release. Fixes go here. |
| 3.x | default branch | Supported until 4.0 leaves beta. Security fixes only after that. |
| 2.x and earlier | - | Not supported. |

## Reporting a vulnerability

**Do not open a public issue for a security problem.**

Use GitHub's private vulnerability reporting: go to the **Security** tab of this
repository and choose **Report a vulnerability**. That opens a private thread
visible only to the maintainer, and it works even though this repository has no
public contact address.

If that is unavailable to you, contact the maintainer through the link in the
repository profile and ask for a private channel before sending any detail.

### What to include

The more of this you can provide, the faster it can be confirmed:

- The version (`4.0.0-beta1`, `3.10.3`, ...) and how you installed it.
- PowerShell edition and version, and the operating system.
- What an attacker can do, not only what looks wrong.
- Steps to reproduce, ideally against a lab tenant.
- Whether it needs an already signed-in session, and what permissions that
  session holds.

**Never include real tenant identifiers, access tokens, client secrets,
certificates or exported policy files from a production tenant.** Redact them, or
reproduce against a lab tenant.

### What to expect

This is a single-maintainer project worked on outside business hours. An
acknowledgement usually takes a few days. A fix ships in the next release for
the affected branch, and the release notes credit the reporter unless you ask
otherwise.

## Scope

This is an administrative client that runs on your own machine with credentials
you supply. Reports that are in scope include:

- Credentials, tokens or secrets written somewhere they should not be, or kept
  in memory or on disk longer than needed.
- A path where the application sends tenant data anywhere other than the Microsoft
  cloud endpoint it is signed in to.
- Code execution from data the application reads: an exported policy file, an
  ADMX file, a documentation template or an imported settings file.
- Anything that causes an operation to run against a different tenant than the
  one selected.

The following are **not** vulnerabilities in this project:

- An account having more permission in Microsoft Intune than you expected. That
  is tenant configuration, not this application.
- Anything requiring an attacker who already controls the machine or the signed-in
  session. At that point they can use the Microsoft Graph API directly.
- Missing hardening that Microsoft Entra or Intune is responsible for, such as
  token lifetime or conditional access.
- Results from a scanner with no demonstrated impact.
