---
name: authorization-expired-functions
description: Interpret SAP authorization analysis results and answer follow-up requests about expired functions, active functions, direct or indirect assignments, or summaries by validity date. Use when the user asks to list only expired items, count expired items, or ask other recurring follow-up questions after an authorization analysis.
---

# Authorization Expired Functions

Use this skill after an authorization analysis has finished and the user asks follow-up questions about the resulting list.

## Expand The Skill

- If the user asks a follow-up pattern that is clearly reusable, suggest adding it to this skill.
- Ask a short confirmation question before treating the new pattern as a permanent rule.
- Prefer adding concrete filters or output formats that recur across analyses, such as:
  - expired only
  - active only
  - direct only
  - indirect only
  - counts by status
  - transactions assigned to a role
  - functions assigned to a user
  - names only / plain list outputs

## Core Rule

- Treat any function or role with `valid_to` earlier than today as expired.
- Treat `31/12/9999` as active.
- If the list already contains a status like `Expirada` or `Ativa`, prefer that status.
- If the user asks for expired functions, return only the expired entries.

## Output Rules

- Keep the answer short and direct.
- Do not repeat active items unless the user explicitly asks for a comparison.
- Preserve the role or function names exactly as shown.
- If the user asks for a count, include the total expired count.
- If the user asks for direct or indirect assignments, use the assignment labels already present in the list.
- If the user asks for transactions assigned to a role, return only the transaction codes or transaction names already present in the analysis results.
- If the user asks for the functions of a role, return the list of functions/TCODEs associated with that role from the final authorization data.
- If the user asks for "a lista das funções", prefer a plain list with one item per line.

## Recommended Filtering

When the final list includes columns or fields like:

- role or function name
- `valid_from`
- `valid_to`
- status
- assignment type

Then:

1. Identify expired items by status `Expirada`, or by `valid_to` before today.
2. Ignore active items.
3. Return the expired names as a plain list.

## Examples

- "Liste-me somente as funcoes expiradas" -> return only expired role or function names.
- "Quantas estao expiradas?" -> return the number plus the expired names if useful.
- "Mostre as expiradas diretas" -> return only entries marked expired and direct.
- "Quais as transações atribuídas à função Z_LOGISTIC_TEMP?" -> return only the transactions linked to that role.
- "Quero a lista das funções" -> return a plain list of the functions/TCODEs from the analyzed result.
