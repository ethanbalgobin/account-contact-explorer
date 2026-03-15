import { useCallback, useEffect, useMemo, useState } from "react";
import type { FormEvent } from "react";
import "./App.css";

import { AccountsService } from "./generated";
import type { Accounts } from "./generated/models/AccountsModel";

import { ContactsService } from "./generated";
import type { Contacts } from "./generated/models/ContactsModel";

type CreateContactInput = Parameters<typeof ContactsService.create>[0];
type UpdateContactInput = Parameters<typeof ContactsService.update>[1];

export default function App() {
  const [accounts, setAccounts] = useState<Accounts[]>([]);
  const [contacts, setContacts] = useState<Contacts[]>([]);
  const [selectedAccountId, setSelectedAccountId] = useState<string | null>(
    null,
  );
  const [selectedContactId, setSelectedContactId] = useState<string | null>(
    null,
  );
  const [accountSearch, setAccountSearch] = useState("");
  const [contactSearch, setContactSearch] = useState("");
  const [newFirstName, setNewFirstName] = useState("");
  const [newLastName, setNewLastName] = useState("");
  const [newEmail, setNewEmail] = useState("");
  const [newJobTitle, setNewJobTitle] = useState("");
  const [newPhone, setNewPhone] = useState("");
  const [isCreatingContact, setIsCreatingContact] = useState(false);
  const [createContactError, setCreateContactError] = useState<string | null>(
    null,
  );
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedUnlinkedContactId, setSelectedUnlinkedContactId] =
    useState<string>("");
  const [isLinkingContact, setIsLinkingContact] = useState(false);
  const [linkContactError, setLinkContactError] = useState<string | null>(null);
  const [contactActionMode, setContactActionMode] = useState<"create" | "link">(
    "create",
  );

  function buildRecordUrl(tableName: string, id: string | null) {
    if (!id) return "#";

    return `${import.meta.env.VITE_DATAVERSE_BASE_URL}/main.aspx?etn=${tableName}&pagetype=entityrecord&id=${encodeURIComponent(
      `{${id}}`,
    )}`;
  }

  function getContactParentAccountId(contact: Contacts): string | null {
    const raw = contact as unknown as Record<string, unknown>;

    const directValue = raw["_parentcustomerid_value"];
    if (typeof directValue === "string" && directValue.length > 0) {
      return directValue;
    }

    const nav = raw["parentcustomerid_account"];
    if (nav && typeof nav === "object") {
      const navObj = nav as Record<string, unknown>;
      if (typeof navObj.accountid === "string" && navObj.accountid.length > 0) {
        return navObj.accountid;
      }
    }

    return null;
  }

  const loadData = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);

      const [accountsResult, contactsResult] = await Promise.all([
        AccountsService.getAll({
          select: [
            "accountid",
            "name",
            "accountnumber",
            "telephone1",
            "address1_city",
          ],
          orderBy: ["name asc"],
          top: 50,
        }),
        ContactsService.getAll({
          select: [
            "contactid",
            "fullname",
            "firstname",
            "lastname",
            "emailaddress1",
            "telephone1",
            "jobtitle",
            "_parentcustomerid_value",
          ],
          orderBy: ["fullname asc"],
          top: 100,
        }),
      ]);

      const loadedAccounts = accountsResult.data ?? [];
      const loadedContacts = contactsResult.data ?? [];

      setAccounts(loadedAccounts);
      setContacts(loadedContacts);

      setSelectedAccountId((currentSelectedAccountId) => {
        if (currentSelectedAccountId) return currentSelectedAccountId;
        return loadedAccounts[0]?.accountid ?? null;
      });
    } catch (err) {
      console.error(err);
      setError("Failed to load Dataverse data.");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadData();
  }, [loadData]);

  useEffect(() => {
    setCreateContactError(null);
    setLinkContactError(null);
    setSelectedUnlinkedContactId("");
  }, [contactActionMode]);

  const filteredAccounts = useMemo(() => {
    const term = accountSearch.trim().toLowerCase();
    if (!term) return accounts;

    return accounts.filter((a) =>
      `${a.name ?? ""} ${a.accountnumber ?? ""} ${a.address1_city ?? ""} ${
        a.telephone1 ?? ""
      }`
        .toLowerCase()
        .includes(term),
    );
  }, [accounts, accountSearch]);

  const filteredContacts = useMemo(() => {
    if (!selectedAccountId) return [];

    const contactsForSelectedAccount = contacts.filter(
      (contact) => getContactParentAccountId(contact) === selectedAccountId,
    );

    const term = contactSearch.trim().toLowerCase();
    if (!term) return contactsForSelectedAccount;

    return contactsForSelectedAccount.filter((c) =>
      `${c.fullname ?? ""} ${c.emailaddress1 ?? ""} ${c.telephone1 ?? ""} ${
        c.jobtitle ?? ""
      }`
        .toLowerCase()
        .includes(term),
    );
  }, [contacts, selectedAccountId, contactSearch]);

  const unlinkedContacts = useMemo(() => {
    return contacts.filter((contact) => !getContactParentAccountId(contact));
  }, [contacts]);

  useEffect(() => {
    if (filteredContacts.length > 0) {
      setSelectedContactId(filteredContacts[0].contactid ?? null);
    } else {
      setSelectedContactId(null);
    }
  }, [filteredContacts]);

  const selectedAccount =
    filteredAccounts.find((a) => a.accountid === selectedAccountId) ??
    accounts.find((a) => a.accountid === selectedAccountId) ??
    null;

  const selectedContact =
    filteredContacts.find((c) => c.contactid === selectedContactId) ?? null;

  async function handleCreateContact(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();

    if (!selectedAccountId) {
      setCreateContactError("Please select an account first.");
      return;
    }

    if (!newLastName.trim()) {
      setCreateContactError("Last name is required.");
      return;
    }

    try {
      setIsCreatingContact(true);
      setCreateContactError(null);

      const payload = {
        firstname: newFirstName.trim() || undefined,
        lastname: newLastName.trim(),
        emailaddress1: newEmail.trim() || undefined,
        jobtitle: newJobTitle.trim() || undefined,
        telephone1: newPhone.trim() || undefined,
        "parentcustomerid_account@odata.bind": `accounts(${selectedAccountId})`,
      };

      const result = await ContactsService.create(
        payload as unknown as CreateContactInput,
      );

      const createdContactId =
        ((result?.data as unknown as Record<string, unknown> | undefined)
          ?.contactid as string | null | undefined) ?? null;

      setNewFirstName("");
      setNewLastName("");
      setNewEmail("");
      setNewJobTitle("");
      setNewPhone("");

      await loadData();

      if (typeof createdContactId === "string" && createdContactId.length > 0) {
        setSelectedContactId(createdContactId);
      }
    } catch (err) {
      console.error("Failed to create contact:", err);
      setCreateContactError("Failed to create contact.");
    } finally {
      setIsCreatingContact(false);
    }
  }

  async function handleLinkExistingContact(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();

    if (!selectedAccountId) {
      setLinkContactError("Please select an account first");
      return;
    }

    if (!selectedUnlinkedContactId) {
      setLinkContactError("Please choose a contact to link");
      return;
    }

    try {
      setIsLinkingContact(true);
      setLinkContactError(null);

      const changes = {
        "parentcustomerid_account@odata.bind": `accounts(${selectedAccountId})`,
      };

      await ContactsService.update(
        selectedUnlinkedContactId,
        changes as unknown as UpdateContactInput,
      );

      await loadData();
      setSelectedContactId(selectedUnlinkedContactId);
      setSelectedUnlinkedContactId("");
    } catch (err) {
      console.error("Failed to link contact to account:", err);
      setLinkContactError("Failed ot link contact to the selected account");
    } finally {
      setIsLinkingContact(false);
    }
  }

  if (loading) {
    return (
      <div className="page">
        <p>Loading Dataverse sample data...</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className="page">
        <h1>Account + Contact Explorer</h1>
        <p className="error">{error}</p>
      </div>
    );
  }

  return (
    <div className="page">
      <header className="hero">
        <h1>Account + Contact Explorer</h1>
        <p>Built as a Power Apps code app against Dataverse sample data.</p>
      </header>

      <div className="layout">
        <section className="card">
          <div className="sectionHeader">
            <h2>Accounts</h2>
            <input
              value={accountSearch}
              onChange={(e) => setAccountSearch(e.target.value)}
              placeholder="Search accounts"
            />
          </div>

          <div className="list">
            {filteredAccounts.map((account) => (
              <button
                key={account.accountid}
                type="button"
                className={`listItem ${
                  selectedAccountId === account.accountid ? "active" : ""
                }`}
                onClick={() => setSelectedAccountId(account.accountid ?? null)}
              >
                <strong>{account.name ?? "(No name)"}</strong>
                <span>{account.accountnumber ?? "No account number"}</span>
                <span>{account.address1_city ?? "No city"}</span>
              </button>
            ))}

            {filteredAccounts.length === 0 && <p>No matching accounts.</p>}
          </div>
        </section>

        <section className="card">
          <h2>Selected Account</h2>

          {selectedAccount ? (
            <div className="details">
              <div>
                <strong>Name:</strong> {selectedAccount.name ?? "—"}
              </div>
              <div>
                <strong>Account Number:</strong>{" "}
                {selectedAccount.accountnumber ?? "—"}
              </div>
              <div>
                <strong>City:</strong> {selectedAccount.address1_city ?? "—"}
              </div>
              <div>
                <strong>Phone:</strong> {selectedAccount.telephone1 ?? "—"}
              </div>
              <div>
                <strong>ID:</strong> {selectedAccount.accountid ?? "—"}
              </div>

              {selectedAccount.accountid && (
                <a
                  className="recordLink"
                  href={buildRecordUrl("account", selectedAccount.accountid)}
                  target="_blank"
                  rel="noreferrer"
                >
                  Open account record
                </a>
              )}
            </div>
          ) : (
            <p>Select an account from the left.</p>
          )}
        </section>
      </div>

      <section className="card contactsCard">
        <div className="sectionHeader">
          <h2>Manage Contact</h2>
          <div className="segmentedControl">
            <button
              type="button"
              className={`segmentButton ${contactActionMode === "create" ? "active" : ""}`}
              onClick={() => setContactActionMode("create")}
            >
              Create new
            </button>
            <button
              type="button"
              className={`segmentButton ${contactActionMode === "link" ? "active" : ""}`}
              onClick={() => setContactActionMode("link")}
            >
              Link existing
            </button>
          </div>
        </div>

        {selectedAccount ? (
          <>
            <p>
              For <strong>{selectedAccount.name ?? "selected account"}</strong>
            </p>

            {contactActionMode === "create" ? (
              <form className="contactForm" onSubmit={handleCreateContact}>
                <input
                  value={newFirstName}
                  onChange={(e) => setNewFirstName(e.target.value)}
                  placeholder="First name"
                />

                <input
                  value={newLastName}
                  onChange={(e) => setNewLastName(e.target.value)}
                  placeholder="Last name *"
                  required
                />

                <input
                  value={newEmail}
                  onChange={(e) => setNewEmail(e.target.value)}
                  placeholder="Email"
                  type="email"
                />

                <input
                  value={newJobTitle}
                  onChange={(e) => setNewJobTitle(e.target.value)}
                  placeholder="Job title"
                />

                <input
                  value={newPhone}
                  onChange={(e) => setNewPhone(e.target.value)}
                  placeholder="Phone"
                />

                <button type="submit" disabled={isCreatingContact}>
                  {isCreatingContact ? "Creating..." : "Create contact"}
                </button>

                {createContactError && (
                  <p className="error">{createContactError}</p>
                )}
              </form>
            ) : (
              <form
                className="contactForm"
                onSubmit={handleLinkExistingContact}
              >
                <div className="selectWrapper">
                  <select
                    className="contactSelect"
                    value={selectedUnlinkedContactId}
                    onChange={(e) =>
                      setSelectedUnlinkedContactId(e.target.value)
                    }
                  >
                    <option value="">Select a contact</option>
                    {unlinkedContacts.map((contact) => (
                      <option
                        key={contact.contactid}
                        value={contact.contactid ?? ""}
                      >
                        {contact.fullname ?? "(No full name)"}
                        {contact.emailaddress1
                          ? ` — ${contact.emailaddress1}`
                          : ""}
                      </option>
                    ))}
                  </select>
                </div>

                <button
                  type="submit"
                  disabled={isLinkingContact || unlinkedContacts.length === 0}
                >
                  {isLinkingContact ? "Linking..." : "Link to selected account"}
                </button>

                {unlinkedContacts.length === 0 && (
                  <p>No unlinked contacts are currently loaded.</p>
                )}

                {linkContactError && (
                  <p className="error">{linkContactError}</p>
                )}
              </form>
            )}
          </>
        ) : (
          <p>Select an account first.</p>
        )}
      </section>

      <section className="card contactsCard">
        <div className="sectionHeader">
          <h2>
            Contacts for {selectedAccount?.name ?? "selected account"} (
            {filteredContacts.length})
          </h2>
          <input
            value={contactSearch}
            onChange={(e) => setContactSearch(e.target.value)}
            placeholder="Search contacts"
          />
        </div>

        <div className="contactsGrid">
          {filteredContacts.map((contact) => (
            <button
              key={contact.contactid}
              type="button"
              className={`contactTile ${
                selectedContactId === contact.contactid ? "active" : ""
              }`}
              onClick={() => setSelectedContactId(contact.contactid ?? null)}
            >
              <strong>{contact.fullname ?? "(No full name)"}</strong>
              <span>{contact.jobtitle ?? "No job title"}</span>
              <span>{contact.emailaddress1 ?? "No email"}</span>
              <span>{contact.telephone1 ?? "No phone"}</span>
            </button>
          ))}

          {filteredContacts.length === 0 && (
            <p>
              {selectedAccount
                ? "No contacts found for this account."
                : "Select an account to see related contacts."}
            </p>
          )}
        </div>
      </section>

      <section className="card contactsCard">
        <h2>Selected Contact</h2>

        {selectedContact ? (
          <div className="details">
            <div>
              <strong>Full Name:</strong> {selectedContact.fullname ?? "—"}
            </div>
            <div>
              <strong>Job Title:</strong> {selectedContact.jobtitle ?? "—"}
            </div>
            <div>
              <strong>Email:</strong> {selectedContact.emailaddress1 ?? "—"}
            </div>
            <div>
              <strong>Phone:</strong> {selectedContact.telephone1 ?? "—"}
            </div>
            <div>
              <strong>ID:</strong> {selectedContact.contactid ?? "—"}
            </div>

            {selectedContact.contactid && (
              <a
                className="recordLink"
                href={buildRecordUrl("contact", selectedContact.contactid)}
                target="_blank"
                rel="noreferrer"
              >
                Open contact record
              </a>
            )}
          </div>
        ) : (
          <p>No contact selected.</p>
        )}
      </section>
    </div>
  );
}
