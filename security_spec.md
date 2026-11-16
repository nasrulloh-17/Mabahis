# Security Specification & Payload-First TDD

## 1. Data Invariants
1. A registration submission (`/registrations/{registrationId}`) MUST reference either an active Firestore form (`/forms/{formId}` where `isActive == true`) or one of the 4 built-in catalog form IDs when not deactivated in Firestore.
2. Public visitors can create a registration submission only with `status == 'menunggu'`, exact schema keys (`hasAll` + `hasOnly`), bounded string sizes, and server-verified timestamps (`createdAt == request.time`, `updatedAt == request.time`).
3. PII (`fullName`, `email`, `phone`, `organization`, `notes`, `customAnswers`) in `/registrations/{registrationId}` is strictly isolated: `get` and `list` are permitted ONLY to `isAdmin()` or the verified owner (`submitterUid == request.auth.uid` with `email_verified == true`).
4. Only verified administrators (`isAdmin()`) can create, update, or delete forms in `/forms/{formId}`, or update/delete registration records.
5. `createdAt`, `id`, `registrationCode`, `formId`, and `submitterUid` are immutable after creation.

## 2. Dirty Dozen Adversarial Payloads (Expected: PERMISSION_DENIED)
1. **Shadow Field Injection on Create**: `{ ...validRegistration, isAdminApproved: true }` -> Blocked by `hasOnly`.
2. **Self-Approved Status on Create**: `{ ...validRegistration, status: 'diterima' }` -> Blocked by `status == 'menunggu'`.
3. **UID Spoofing on Create**: Unauthenticated or mismatched `submitterUid: 'victim-uid-123'` -> Blocked by `submitterUid` identity gate.
4. **Orphan / Inactive Form Reference**: `{ ...validRegistration, formId: 'non-existent-form' }` -> Blocked by `exists()` + `isActive == true`.
5. **1MB String Exhaustion Attack**: `{ ...validRegistration, notes: 'x'.repeat(20000) }` -> Blocked by `notes.size() <= 1000`.
6. **ID Poisoning Attack**: Document ID `'invalid id with spaces and $$$'` -> Blocked by `isValidId()`.
7. **Timestamp Forgery**: `{ ...validRegistration, createdAt: Timestamp(2020) }` -> Blocked by `createdAt == request.time`.
8. **Blanket List Scraping by Non-Admin**: `getDocs(collection(db, 'registrations'))` without `where('submitterUid', '==', uid)` -> Blocked by `allow list`.
9. **Cross-User PII Read**: Signed-in user A reading `/registrations/userB-doc` -> Blocked by `existing().submitterUid == request.auth.uid`.
10. **Non-Admin Form Creation**: Signed-in regular user writing to `/forms/malicious-form` -> Blocked by `isAdmin()`.
11. **Non-Admin Privilege Escalation**: Signed-in regular user writing `/admins/{uid}` -> Blocked by `allow read, write: if false`.
12. **Unverified Email Admin Spoof**: Token with admin email string but `email_verified: false` -> Blocked by `request.auth.token.email_verified == true`.
