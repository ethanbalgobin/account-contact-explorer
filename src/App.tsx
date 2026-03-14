import { useEffect, useMemo, useState } from "react";
import "./App.css";

import { AccountsService } from "./generated";
import type { Accounts } from "./generated/models/AccountsModel";

import { ContactsService } from "./generated";
import type { Contacts } from "./generated/models/ContactsModel";

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

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  function buildRecordUrl(tableName: string, id: string | null) {
    if (!id) return "#";

    return `${import.meta.env.VITE_DATAVERSE_BASE_URL}/main.aspx?etn=${tableName}&pagetype=entityrecord&id=${encodeURIComponent(`{${id}}`)}`;
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

  useEffect(() => {
    const load = async () => {
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
              "emailaddress1",
              "telephone1",
              "jobtitle",
              "_parentcustomerid_value",
            ],
            orderBy: ["fullname asc"],
            top: 50,
          }),
        ]);

        console.log("Contacts result:", contactsResult);
        console.log("First contact sample:", contactsResult.data?.[0]);

        const loadedAccounts = accountsResult.data ?? [];
        const loadedContacts = contactsResult.data ?? [];

        setAccounts(loadedAccounts);
        setContacts(loadedContacts);

        if (loadedAccounts.length > 0 && loadedAccounts[0].accountid) {
          setSelectedAccountId(loadedAccounts[0].accountid);
        }
      } catch (err) {
        console.error(err);
        setError("Failed to load Dataverse data.");
      } finally {
        setLoading(false);
      }
    };

    load();
  }, []);

  const filteredAccounts = useMemo(() => {
    const term = accountSearch.trim().toLowerCase();
    if (!term) return accounts;

    return accounts.filter((a) =>
      `${a.name ?? ""} ${a.accountnumber ?? ""} ${a.address1_city ?? ""} ${a.telephone1 ?? ""}`
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
      `${c.fullname ?? ""} ${c.emailaddress1 ?? ""} ${c.telephone1 ?? ""} ${c.jobtitle ?? ""}`
        .toLowerCase()
        .includes(term),
    );
  }, [contacts, selectedAccountId, contactSearch]);

  useEffect(() => {
    if (filteredContacts.length > 0) {
      setSelectedContactId(filteredContacts[0].contactid ?? null);
    } else {
      setSelectedAccountId(null);
    }
  }, [filteredContacts]);

  const selectedAccount =
    filteredAccounts.find((a) => a.accountid === selectedAccountId) ??
    accounts.find((a) => a.accountid === selectedAccountId) ??
    null;

  const selectedContact =
    filteredContacts.find((c) => c.contactid === selectedContactId) ?? null;

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
                className={`listItem ${selectedAccountId === account.accountid ? "active" : ""}`}
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
              {selectedAccount?.accountid && (
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
              className={`contactTile ${selectedContactId === contact.contactid ? "active" : ""}`}
              onClick={() => setSelectedContactId(contact.contactid ?? null)}
              type="button"
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
      </section>
    </div>
  );
}
