// Node 18 lacks the Web `File` global that Node 20+ and browsers provide.
// The generated Graph client uses `z.instanceof(File)` for multipart upload
// endpoints, so importing it under Node 18 throws ReferenceError at module load.
// Provide a minimal stand-in so tests that load the real client can run on 18.
if (typeof (globalThis as { File?: unknown }).File === 'undefined') {
  (globalThis as { File?: unknown }).File = class File {};
}

// The server-side send gate (src/send-gate.ts) refuses every send when GATE_CHECK_URL
// is unset. Suites that exercise send tools for other reasons run with it disabled;
// test/send-gate.test.ts sets the gate env per test.
if (process.env.GATE_DISABLED === undefined) process.env.GATE_DISABLED = 'true';
