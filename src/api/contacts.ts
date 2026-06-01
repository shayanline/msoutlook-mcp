/**
 * Contacts API.
 */

import { owaGet, owaPost, owaPatch, owaDelete } from './client.js';
import type { ODataResponse } from './mail.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

export interface Contact {
  Id: string;
  DisplayName: string;
  GivenName?: string;
  Surname?: string;
  EmailAddresses?: Array<{ Name?: string; Address: string }>;
  BusinessPhones?: string[];
  MobilePhone?: string;
  JobTitle?: string;
  CompanyName?: string;
  Department?: string;
  OfficeLocation?: string;
  BusinessAddress?: {
    Street?: string;
    City?: string;
    State?: string;
    CountryOrRegion?: string;
    PostalCode?: string;
  };
  Notes?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// List contacts
// ─────────────────────────────────────────────────────────────────────────────

export async function listContacts(search?: string, top = 25): Promise<Contact[]> {
  const params: Record<string, string> = {
    '$top': String(top),
    '$orderby': 'displayName asc',
    '$select': 'Id,DisplayName,GivenName,Surname,EmailAddresses,BusinessPhones,MobilePhone,JobTitle,CompanyName',
  };

  if (search) params['$search'] = `"${search}"`;

  const res = await owaGet<ODataResponse<Contact>>('/contacts', params);
  return res.value;
}

// ─────────────────────────────────────────────────────────────────────────────
// Get contact
// ─────────────────────────────────────────────────────────────────────────────

export async function getContact(id: string): Promise<Contact> {
  return owaGet<Contact>(`/contacts/${id}`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Create contact
// ─────────────────────────────────────────────────────────────────────────────

export interface CreateContactOptions {
  givenName?: string;
  surname?: string;
  displayName?: string;
  email?: string;
  businessPhone?: string;
  mobilePhone?: string;
  jobTitle?: string;
  companyName?: string;
  notes?: string;
}

export async function createContact(opts: CreateContactOptions): Promise<Contact> {
  const emailAddresses = opts.email ? [{ Address: opts.email }] : undefined;

  return owaPost<Contact>('/contacts', {
    GivenName: opts.givenName,
    Surname: opts.surname,
    DisplayName: opts.displayName ?? `${opts.givenName ?? ''} ${opts.surname ?? ''}`.trim(),
    ...(emailAddresses ? { EmailAddresses: emailAddresses } : {}),
    ...(opts.businessPhone ? { BusinessPhones: [opts.businessPhone] } : {}),
    MobilePhone: opts.mobilePhone,
    JobTitle: opts.jobTitle,
    CompanyName: opts.companyName,
    PersonalNotes: opts.notes,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Delete contact
// ─────────────────────────────────────────────────────────────────────────────

export async function deleteContact(id: string): Promise<void> {
  await owaDelete(`/contacts/${id}`);
}
