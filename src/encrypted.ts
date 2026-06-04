// Detect a password-encrypted .docx. Encrypted Office files are an OLE/CFB compound document
// (magic D0 CF 11 E0 A1 B1 1A E1), not a ZIP — opening them with jszip fails confusingly, so we
// detect early and surface a clear message. (Agile AES decryption, as in SuperDoc's agile-decryptor,
// is a larger follow-up.)
const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export function isEncryptedOfficeFile(bytes: Uint8Array): boolean {
  if (bytes.length < OLE_MAGIC.length) return false;
  return OLE_MAGIC.every((b, i) => bytes[i] === b);
}
