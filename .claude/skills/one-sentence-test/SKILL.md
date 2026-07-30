---
name: one-sentence-test
description: Force one concrete sentence to find orphaned surfaces, duplicate verbs, inert abstractions. Use when evaluating design coherence or whether an abstraction earns its keep.
author: epicenter
version: "2.0"
---

Before continuing, stop and write one concrete sentence that describes the subject. Name the objects, the verbs, and the scope. No marketing language. Write the sentence visibly — readers will see the gap between prose claims and reality. If the sentence drifts, the design isn't coherent yet.

## Application A — Cohesion Audit (Top-Down)

**Subject:** A design, spec, or surface (commands, endpoints, tables) whose coherence is in question. The sentence becomes the thesis; every surface is audited against it.

**Signals:** Designs wrapping up, specs with unrelated sections, growing command lists, contributors describing the system differently, users asking "what is this" or "elevator pitch."

**Audit questions:**
- Does each surface serve the sentence?
- If we removed this surface, would the sentence still be true?
- Is there a verb in the sentence with no surface?
- Do two surfaces redundantly serve the same verb?
- Does a convenience feature force a second product sentence?

**Asymmetric wins check (4-step):**
1. List conveniences
2. Circle the one creating the most surface area
3. Remove it from the sentence
4. If it's still useful, run the asymmetric-wins pass

**Worked example:** A notification API (`notify.send`, `notify.schedule`, etc.) starts as "A notification system" (too vague) → thesis: "Deliver a message to a user on their preferred channel, honoring their quiet hours, and remember we sent it." Audit finds `schedule`, `batch`, and `digest` don't fit — they belong to a sibling product.

**Good vs. Bad sentences:**
- Bad: "A CLI for managing your workspace"
- Good: Names specific objects (actions in `epicenter.config.ts`), verbs (introspect, invoke), and scope (local or live peer)

## Application B — Value-Add Audit (Bottom-Up)

**Subject:** A single utility, wrapper, flag, endpoint, or config option whose value is in question. The sentence describes what code *actually does*, ignoring docs, then specializes under active defaults.

**Signals:** Evaluating whether an abstraction earns its keep, reviewing wrappers, writing docs, when code and docs disagree.

**Three required reductions:**

1. **Strip-docs reduction** — Ignore name/JSDoc/README, read the body, state in plain language what the function does. No adjectives. Just the mechanics.
2. **Default-config reduction** — Specialize under defaults actually in use; features behind non-default options don't count.
3. **Caller-need check** — Does the caller's use case actually need the default-config version?

**Worked example:** `createDisposableCache` with extensive JSDoc. Strip-docs: refcounting, dedup, idle timeout, eviction. Default-config (gcTime: Infinity): "Close is explicit: docs live forever" — refcount is inert, no eviction fires. Caller wraps a singleton with one id and one lifetime, so dedup is a no-op. Verdict: **Delete it.**

## Common Reveals
- Inert knobs (defaults disabling advertised mechanisms)
- Singleton wrappers
- Ceremony with no payoff
- Docs describing aspirations
- Orphaned surfaces
- Drifting thesis

## Anti-Patterns
- Paraphrasing JSDoc instead of reading the body
- Stopping at strip-docs reduction
- Defending an abstraction by listing unused features
- Calling the thesis "a CLI for workspaces" or similar fluff

## Success Criteria
After the move, you can state in one concrete sentence:
- (Cohesion) What the subject **is**
- (Value-add) What it's **buying this caller**

If you can't, or the honest answer is "nothing, really," the design doesn't earn its keep.

## What This Skill Is Not
- Not a naming exercise (the codename stays)
- Not a marketing tagline (must be concrete)
- Not a complete spec (it's a coherence gate or sniff test)
