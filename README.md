# On-Premises Active Directory & Windows Python Utilities

A collection of Python utilities originally developed to support **on-premises Active Directory** and **Windows endpoint administration**. The scripts use LDAP queries to inspect directory objects and group memberships, plus Windows APIs to review local administrators and installed software.

> **Project status:** Legacy portfolio code. These scripts reflect tools built for a previous Windows/Active Directory environment. They have not been revalidated against a live domain for this publication and may require updates before reuse. Review the [limitations](#known-limitations) and security considerations before running them.

## Highlights

- **Identity and access reviews:** Find a user's AD group memberships and compare memberships between two accounts.
- **Directory inventory:** Locate the distinguished name/OU path of an AD computer account and list direct users in an AD group.
- **Windows administration:** Inspect local Administrators membership and installed software recorded in the Windows registry.
- **Command-line utilities:** Provide arguments using `argparse` and enter an LDAP password interactively using `getpass`.

## Script directory

| Script | Technology | Purpose |
|---|---|---|
| [`group_check.py`](group_check.py) | AD / LDAP | List a user's direct `memberOf` groups, sorted by name. |
| [`match_groups.py`](match_groups.py) | AD / LDAP | Show common direct group memberships for two users. |
| [`diff_groups.py`](diff_groups.py) | AD / LDAP | Show direct groups unique to either of two users. |
| [`list_users_n_group.py`](list_users_n_group.py) | AD / LDAP | List direct **user** members of a named AD group. |
| [`computer_OU.py`](computer_OU.py) | AD / LDAP | Look up the parent distinguished-name path of a computer object. |
| [`list_admin.py`](list_admin.py) | Windows / COM | List local Administrators group members; remote mode is experimental. |
| [`list_installed_apps.py`](list_installed_apps.py) | Windows Registry / Outlook COM | Inventory installed applications and email a text report. |

The repository may also contain development or backup files (such as `hello.py` and `list_users_n_group_copy.py`); these are not considered separate production-ready utilities.

## Script documentation

### Active Directory / LDAP

<details>
<summary><strong>group_check.py</strong> — User group membership lookup</summary>

- **Input:** One AD username (`sAMAccountName`).
- **Output:** Sorted names of groups listed in the user's `memberOf` attribute.
- **Example:** `python group_check.py jdoe`
- **Use case:** Check a user's assigned AD groups during access reviews or troubleshooting.
- **Limitations:** Direct `memberOf` memberships only; no nested group expansion or complete effective group-token calculation. The script assumes the user is found.

</details>

<details>
<summary><strong>match_groups.py</strong> — Compare common user groups</summary>

- **Input:** Two AD usernames.
- **Output:** Common group names found by intersecting the users' direct `memberOf` groups.
- **Example:** `python match_groups.py jdoe asmith`
- **Use case:** Compare access group assignments when troubleshooting or preparing a new user's access.
- **Limitations:** Uses simplified CN strings rather than full distinguished names, does not expand nested memberships, and assumes both accounts exist.

</details>

<details>
<summary><strong>diff_groups.py</strong> — Compare different user groups</summary>

- **Input:** Two AD usernames.
- **Output:** Groups unique to the first user and groups unique to the second user.
- **Example:** `python diff_groups.py jdoe asmith`
- **Use case:** Identify potential differences in group-based access assignments.
- **Limitations:** Direct memberships only; simplified group-name comparisons can be ambiguous if names repeat in different OUs.

</details>

<details>
<summary><strong>list_users_n_group.py</strong> — Enumerate group users</summary>

- **Input:** A group's exact CN.
- **Output:** Direct user members' `sAMAccountName` and display name, sorted by display name.
- **Example:** `python list_users_n_group.py "Accounting Team"`
- **Use case:** Review directly assigned users in an AD group.
- **Limitations:** Resolves person objects, not nested group membership; users in nested groups are not expanded.

</details>

<details>
<summary><strong>computer_OU.py</strong> — Computer directory location</summary>

- **Input:** AD computer name (without the trailing `$`).
- **Output:** Parent distinguished-name path, or a not-found message.
- **Example:** `python computer_OU.py WORKSTATION01`
- **Use case:** Locate computer objects for organizational-unit administration and troubleshooting.
- **Limitations:** Extracts the parent DN with a simple comma split, which may not correctly handle escaped commas in distinguished names.

</details>

### Windows administration

<details>
<summary><strong>list_admin.py</strong> — Local administrator membership</summary>

- **Input:** Optional computer name and optional `--username` argument.
- **Output:** Names returned as members of the local Administrators group.
- **Example (local):** `python list_admin.py`
- **Use case:** Inspect local administrative access.
- **Limitations:** Windows-only (`pywin32`). The original source notes that remote-computer queries were not working reliably; **do not rely on remote results without fixing and testing that path**. Output is printed to the console only.

</details>

<details>
<summary><strong>list_installed_apps.py</strong> — Installed software report</summary>

- **Input:** No command-line arguments; requires an `email` setting in `config.ini` and a configured Outlook desktop client.
- **Output:** Console list, `installed_apps.txt` in the working directory, and an Outlook message with that file attached.
- **Use case:** Collect a local Windows software inventory for support and troubleshooting.
- **Limitations:** Windows-only (`winreg` and `pywin32`); reads selected `HKLM` uninstall registry paths, not every possible installation source or per-user application. **Running the script automatically sends an email** to the configured recipient; verify the recipient and contents first.

</details>

## Requirements

- Python 3
- Network connectivity and authorized read access to the on-premises Active Directory server for LDAP scripts
- `ldap3` for the Active Directory scripts
- Windows and `pywin32` for the Windows COM scripts
- Microsoft Outlook desktop configured for the installed-apps email workflow
- An environment-specific `config.ini` file for the LDAP utilities and installed-apps script

Install the third-party dependencies where needed:

```powershell
python -m pip install ldap3 pywin32
```

`pywin32` is only needed on Windows for the two COM-based utilities.

## Configuration

The LDAP scripts expect a `config.ini` file in the **current working directory**, with keys similar to:

```ini
[DEFAULT]
server = ad.example.com
user_dn = CN=Audit User,OU=Service Accounts,DC=example,DC=com
base_dn = DC=example,DC=com
email = admin@example.com
```

- `server`: Directory server address.
- `user_dn`: Distinguished name used to bind to LDAP.
- `base_dn`: Search base for directory queries.
- `email`: Recipient used by `list_installed_apps.py` (not needed by the other scripts).

This is an illustrative configuration, **not a working connection**. The scripts prompt for an LDAP password where needed. Store environment-specific configuration privately; never commit actual credentials or internal directory details.

**Transport security:** Passwords are entered without terminal echo, but that does not itself encrypt LDAP traffic. Use an appropriately secured LDAP connection (LDAPS or StartTLS with certificate verification) before using the utilities with real credentials.

## Example usage

From a terminal in the repository directory:

```powershell
python group_check.py jdoe
python match_groups.py jdoe asmith
python diff_groups.py jdoe asmith
python list_users_n_group.py "Accounting Team"
python computer_OU.py WORKSTATION01
python list_admin.py
```

For the installed-apps report, review the configured recipient before running:

```powershell
python list_installed_apps.py
```

## Known limitations

- **LDAP input handling:** LDAP search values are currently interpolated directly into filters. User-supplied values should be escaped using `ldap3.utils.conv.escape_filter_chars` before production use.
- **Missing-object handling:** Some scripts assume search results exist and can raise errors when users or groups are missing.
- **Group scope:** The group-comparison scripts use `memberOf` and do not recursively expand nested memberships or calculate all effective access.
- **DN parsing:** Some scripts simplify distinguished names by splitting strings; a proper DN parser is safer.
- **Output:** Most scripts print to the console rather than export structured CSV reports.
- **Remote administrator inventory:** The remote mode of `list_admin.py` is not confirmed working.
- **Software inventory:** `list_installed_apps.py` reads selected system registry uninstall locations and automatically sends an Outlook email.

## Security and publishing

Before using or publishing this repository:

1. Confirm you have permission to publish code developed during prior employment.
2. Remove employer-specific hostnames, email addresses, credentials, user names, and output data from code, configuration, and Git history.
3. Keep `config.ini`, generated reports, and local environment files out of source control.
4. Use encrypted and authenticated LDAP transport with least-privilege directory access.
5. Test any refactored versions in a lab or other authorized Active Directory environment.

Suggested `.gitignore` additions:

```gitignore
config.ini
*.csv
installed_apps.txt
__pycache__/
*.py[cod]
.venv/
.env
```

## About

These are **legacy administration scripts** illustrating Python-based LDAP queries and Windows support automation. They complement newer projects focused on Microsoft 365, Microsoft Entra ID, and Google Workspace APIs by demonstrating experience with traditional on-premises identity infrastructure.
