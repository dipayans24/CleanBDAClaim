// app/page.tsx
// The entire user interface: choose a CSV/Excel file, upload it to
// /api/process, and download whatever comes back. This is a client
// component (not server-rendered) because it needs onChange/onClick handlers.
"use client";

import { useState } from "react";
import type { ChangeEvent, FormEvent } from "react";

export default function HomePage() {
  const [file, setFile] = useState<File | null>(null); // the file the user picked
  const [isProcessing, setIsProcessing] = useState(false); // true while the request is in flight
  const [statusMessage, setStatusMessage] = useState<string | null>(null); // success text
  const [errorMessage, setErrorMessage] = useState<string | null>(null); // failure text

  // Runs when the user picks (or clears) a file in the <input type="file">
  function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const selected = event.target.files?.[0] ?? null; // the browser only ever gives us one file here
    setFile(selected);
    setErrorMessage(null); // clear any previous error once a new file is chosen
    setStatusMessage(null);
  }

  // Runs when the form is submitted: sends the file to the API and downloads the result
  async function handleSubmit(event: FormEvent) {
    event.preventDefault(); // stop the browser's default full-page form submit

    if (!file) {
      setErrorMessage("Choose a CSV or Excel file first.");
      return;
    }

    setIsProcessing(true);
    setErrorMessage(null);
    setStatusMessage(null);

    try {
      const formData = new FormData(); // multipart body, matches what the API route expects
      formData.append("file", file);

      const response = await fetch("/api/process", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        // On failure the API returns JSON like { error: "..." }
        const body = await response.json().catch(() => ({}));
        throw new Error(body.error ?? "Processing failed.");
      }

      // On success, the response body IS the cleaned .xlsx file's bytes
      const blob = await response.blob();

      // Pull the filename the server chose out of the Content-Disposition header
      const disposition = response.headers.get("Content-Disposition") ?? "";
      const match = disposition.match(/filename="?([^"]+)"?/);
      const downloadName = match ? match[1] : "cleaned.xlsx";

      // Trigger a normal browser download via a temporary, invisible link
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = downloadName;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url); // release the memory now that the download has started

      setStatusMessage(`Done — ${downloadName} has been downloaded.`);
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : "Something went wrong.");
    } finally {
      setIsProcessing(false);
    }
  }

  return (
    <main className="page">
      <header className="masthead">
        <p className="kicker">Workshop &amp; payment exports</p>
        <h1>BDA file cleaner</h1>
        <p>
          Upload the raw CSV or Excel export from LeadSquared. It runs through the same cleaning
          rules as the original script — dates parsed, blank phone numbers filled, refunds split
          out, final payment computed — and the cleaned workbook downloads automatically.
        </p>
      </header>

      <form className="dropzone" onSubmit={handleSubmit}>
        <label htmlFor="bda-file">CSV or Excel file (.csv, .xlsx, .xls)</label>
        <div className="file-row">
          <input id="bda-file" type="file" accept=".csv,.xlsx,.xls" onChange={handleFileChange} />
          <button type="submit" disabled={isProcessing}>
            {isProcessing ? "Processing…" : "Upload & clean"}
          </button>
        </div>
        {file && <p className="file-name">Selected: {file.name}</p>}
      </form>

      {errorMessage && <p className="status error">{errorMessage}</p>}
      {statusMessage && <p className="status success">{statusMessage}</p>}

      <section className="rules">
        <h2>What this does to your file</h2>
        <ol>
          <li>
            Parses <code>createdAt</code>, <code>updatedAt</code> and <code>date_time_of_closing</code> as{" "}
            <code>DD/MM/YYYY HH:mm</code>.
          </li>
          <li>
            Fills any <code>phoneNumber</code> of <code>0</code> from <code>original_phone_number</code>, then
            assigns sequential placeholders (starting at 101) to whatever is still blank.
          </li>
          <li>
            Splits out <code>partial_refunded</code> rows and computes their final payment as amount paid
            minus amount refunded, then drops every other refunded row.
          </li>
          <li>
            Inserts a <code>FinalPayment</code> column before <code>paymentAmount</code>, using the larger of{" "}
            <code>closing_amt</code>/<code>rest_collection</code> when <code>paymentAmount</code> is 0.
          </li>
          <li>
            Blanks out <code>alternateEmails</code>/<code>alternatePhoneNumbers</code> values 71 characters or
            longer, and drops the <code>paymentproof</code> and <code>remark</code> columns.
          </li>
        </ol>
      </section>
    </main>
  );
}
