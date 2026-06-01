/**
 * Contacts and people MCP tools.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { listContacts, getContact, createContact, deleteContact, type Contact } from '../api/contacts.js';
import { searchPeople, type Person } from '../api/people.js';

// ─────────────────────────────────────────────────────────────────────────────
// Formatters
// ─────────────────────────────────────────────────────────────────────────────

function formatContact(c: Contact): string {
  const emails = c.EmailAddresses?.map(e => e.Address).join(', ') ?? '';
  const phones = [...(c.BusinessPhones ?? []), c.MobilePhone].filter(Boolean).join(', ');
  return [
    `ID: ${c.Id}`,
    `Name: ${c.DisplayName}`,
    emails ? `Email: ${emails}` : '',
    phones ? `Phone: ${phones}` : '',
    c.JobTitle ? `Title: ${c.JobTitle}` : '',
    c.CompanyName ? `Company: ${c.CompanyName}` : '',
    c.Department ? `Department: ${c.Department}` : '',
  ].filter(Boolean).join('\n');
}

function formatPerson(p: Person): string {
  const email = p.ScoredEmailAddresses?.[0]?.Address
    ?? p.EmailAddresses?.[0]?.Address
    ?? '';
  return [
    `Name: ${p.DisplayName}`,
    email ? `Email: ${email}` : '',
    p.UserPrincipalName ? `UPN: ${p.UserPrincipalName}` : '',
    p.JobTitle ? `Title: ${p.JobTitle}` : '',
    p.Department ? `Department: ${p.Department}` : '',
    p.CompanyName ? `Company: ${p.CompanyName}` : '',
  ].filter(Boolean).join('\n');
}

// ─────────────────────────────────────────────────────────────────────────────
// Tools
// ─────────────────────────────────────────────────────────────────────────────

export function registerContactTools(server: McpServer): void {
  // ── outlook_list_contacts ────────────────────────────────────────────────
  server.tool(
    'outlook_list_contacts',
    'List contacts from the Outlook address book.',
    {
      search: z.string().optional().describe('Search by name or email'),
      top: z.number().int().min(1).max(100).optional().describe('Max results (default 25)'),
    },
    async ({ search, top }) => {
      const contacts = await listContacts(search, top ?? 25);
      if (contacts.length === 0) return { content: [{ type: 'text', text: 'No contacts found.' }] };
      const text = contacts.map((c, i) => `--- Contact ${i + 1} ---\n${formatContact(c)}`).join('\n\n');
      return { content: [{ type: 'text', text }] };
    },
  );

  // ── outlook_get_contact ──────────────────────────────────────────────────
  server.tool(
    'outlook_get_contact',
    'Get full details of a contact by ID.',
    {
      id: z.string().describe('Contact ID'),
    },
    async ({ id }) => {
      const contact = await getContact(id);
      return { content: [{ type: 'text', text: formatContact(contact) }] };
    },
  );

  // ── outlook_create_contact ───────────────────────────────────────────────
  server.tool(
    'outlook_create_contact',
    'Create a new contact in Outlook.',
    {
      given_name: z.string().optional().describe('First name'),
      surname: z.string().optional().describe('Last name'),
      email: z.string().email().optional().describe('Email address'),
      business_phone: z.string().optional().describe('Business phone number'),
      mobile_phone: z.string().optional().describe('Mobile phone number'),
      job_title: z.string().optional().describe('Job title'),
      company_name: z.string().optional().describe('Company or organisation name'),
      notes: z.string().optional().describe('Personal notes'),
    },
    async (params) => {
      const contact = await createContact({
        givenName: params.given_name,
        surname: params.surname,
        email: params.email,
        businessPhone: params.business_phone,
        mobilePhone: params.mobile_phone,
        jobTitle: params.job_title,
        companyName: params.company_name,
        notes: params.notes,
      });
      return {
        content: [{
          type: 'text',
          text: `Contact created.\nID: ${contact.Id}\nName: ${contact.DisplayName}`,
        }],
      };
    },
  );

  // ── outlook_delete_contact ───────────────────────────────────────────────
  server.tool(
    'outlook_delete_contact',
    'Delete a contact by ID.',
    {
      id: z.string().describe('Contact ID to delete'),
    },
    async ({ id }) => {
      await deleteContact(id);
      return { content: [{ type: 'text', text: 'Contact deleted.' }] };
    },
  );

  // ── outlook_search_people ────────────────────────────────────────────────
  server.tool(
    'outlook_search_people',
    'Search the organisation directory for people by name or email.',
    {
      query: z.string().describe('Name, email, or job title to search for'),
      top: z.number().int().min(1).max(25).optional().describe('Max results (default 10)'),
    },
    async ({ query, top }) => {
      const people = await searchPeople(query, top ?? 10);
      if (people.length === 0) return { content: [{ type: 'text', text: 'No people found.' }] };
      const text = people.map((p, i) => `--- Person ${i + 1} ---\n${formatPerson(p)}`).join('\n\n');
      return { content: [{ type: 'text', text }] };
    },
  );
}
