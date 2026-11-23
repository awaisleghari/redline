# Redline

An open-source web app that compares two Word documents (original and modified) and produces a legal-style redline, like Litera Compare. The user can view the redline in the page or download it as Word or PDF. No AI model is used anywhere in the product.

**Read `PLAN.md` at the start of every session.** It holds the architecture, specs, stages, acceptance criteria and the status log.

## Working style

The owner is a corporate/commercial lawyer, not a programmer. He wants this run hands-off: you work for long stretches and he steps in only for important decisions.

- Work through the stages in `PLAN.md` in order, without waiting for approval between ordinary steps.
- Before each stage, write a short plan for it into the status log. After each stage, run the full test suite, run an independent review (a subagent that checks the diff against the stage's acceptance criteria and reports only real gaps), fix what it finds, update the status log, commit and push.
- When you report to the owner, use plain language: what was done, what it means, what is next. No jargon without a one-line explanation.
- Build to the standard of the best engineer you can be. Fix root causes; never silence an error or weaken a test to make it pass.

## Where you are running

Claude Code cloud sessions: a fresh Ubuntu 24.04 virtual machine each session, with Node 22, Python, Docker and git pre-installed. The session keeps working when the owner closes his laptop.

- **Network allowlist.** Only allowlisted domains are reachable. If a download is blocked, do not hunt for workarounds through other hosts: name the exact domain needed and ask the owner to add it to the environment's allowed domains.
- **Git.** Push to the working branch the session gives you, often. At the end of each stage, open or update a pull request to `main` with a plain-English summary. Merge it yourself if you have permission and every check passes; otherwise tell the owner it is ready to merge.
- **Long commands** (Docker builds, full test runs) may outlast the default command timeout. Run them in the background and check on them rather than restarting them.
- **Files outside GitHub do not exist here.** The session sees only what is committed. Git-ignored folders such as `tests/fixtures/private/` are empty in the cloud, by design.

## Decide on your own

Implementation details, file structure, library choice (any open-source license, including GPL/AGPL), test design, bug fixes, refactors, retrying failed approaches, and the order of work inside a stage.

## Stop and ask (never decide these alone)

Ask the owner **in the chat**, using the question tool (multiple choice, your recommended option first, one line on why). He will not read question files. To keep interruptions rare: first finish every piece of work that does not depend on the answer, then ask all pending questions together in one go. Ask when the action:
- costs money or needs a paid plan (Cloudflare Workers Paid, Containers);
- needs the owner's accounts, logins or credentials (Cloudflare, GitHub settings);
- deploys to a public URL for the first time, or changes what the public site does with uploaded files;
- changes scope or the stages in `PLAN.md`;
- rewrites git history or deletes work (never force-push);
- needs real client documents.

Things only a lawyer's eye can judge (does the redline look right in Word, is the styling acceptable) are **owner checks**. Collect them during a stage and ask for them in the chat at the end of the stage, all at once, naming the exact files to open and what to look for. Record the owner's answers in the status log in `PLAN.md`.

## Hard rules

- Never commit client documents, secrets, API keys or `.env` files. Real contracts live only in `tests/fixtures/private/`, which is git-ignored.
- The product is split into a Producer and an independent Checker (see `PLAN.md`). Checker code never imports Producer code, and no output reaches the user unless the Checker passes it.
- Never weaken a Checker test or remove a broken-output case to make something pass. If the Checker and Producer disagree, find which is wrong and fix that one.
- The independent review you run at the end of each stage is a development step. It does not replace the Checker inside the product.
- Documents are confidential. Nothing is sent anywhere except the PDF step described in `PLAN.md`, and nothing about document contents is logged.
- Keep this file short. Details belong in `PLAN.md`.
