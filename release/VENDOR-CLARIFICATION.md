# Prepared vendor questions — not sent

This is a draft for later owner-approved outreach, not vendor confirmation.
No account IDs, credentials, personal paths, business documents or private logs
are needed. Keep any actual correspondence outside the public source snapshot.

## CodeBuddy SDK 0.3.243

Potential contact: `codebuddy@tencent.com`, as published in the official package
README and the [public service agreement, section 4.2](https://www.codebuddy.ai/document/term).
This published customer contact was independently verified on 2026-10-12;
it is not an author/employee account inferred from the domain. No email or issue
has been submitted. See [privacy scope](PRIVACY-REVIEW.md).

Our source-only MIT Office add-in references `@tencent-ai/agent-sdk@0.3.243` but
does not redistribute it. Each user obtains the vendor CLI/SDK separately and
uses their own account. A future optional self-contained desktop package would
require the following clarification:

1. Does the package's MIT declaration cover all files under lib/ and cli/? The
   downloaded package has cli/LICENSE.txt, but no root LICENSE. Which copyright
   and permission notices must accompany the SDK library and CLI respectively?
2. Is redistribution of the unmodified SDK/CLI in a third-party desktop App
   permitted, distinct from users installing it themselves? Are platform/native
   payloads subject to separate grants, and is an exact-version notice/SBOM set
   available for ripgrep, native Node modules, toybox and zsh?
3. How do the SDK MIT grant, SDK documentation's third-party-app use cases and
   service-agreement sections concerning transfer/distribution apply together?
   Which terms govern users of international, China and enterprise environments?
4. May each end user run the unmodified SDK with their separately licensed own
   account through a local Office add-in, without resale/shared credentials or
   a hosted gateway? Are there SDK-specific authentication/usage requirements?
5. What plain-text integration/compatibility wording is permitted? No vendor
   logo, affiliation or official-support claim is proposed.

## Electron internal extract-zip, only if actually bundled

Version 1.0.5 declares BSD-2-Clause at gitHead
b83e459fd04c53b0a1c8438a6792df8f64be47fc but contains no root LICENSE.
[PR #19](https://github.com/electron/extract-zip/pull/19) later changes the metadata
to MIT and adds a license. If we intend to redistribute the pinned installer
module, request an exact-version grant/notice or explicit confirmation of scope
and the compiled-crate notice requirements. Do not request or perform a dependency
upgrade solely to bypass review. If the component is not in the deliverable,
record it as install-only/N/A rather than making outreach a source-release gate.
