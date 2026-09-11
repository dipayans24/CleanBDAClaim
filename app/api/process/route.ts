// app/api/process/route.ts
//
// A single POST endpoint: receives an uploaded CSV/Excel file, runs it
// through the cleaning pipeline in lib/bdaProcessor.ts, and streams back
// the cleaned .xlsx file for the browser to download.

import { NextRequest, NextResponse } from "next/server";
import { readInputFile, processBda, writeOutputWorkbook } from "@/lib/bdaProcessor";

// The xlsx library needs Node's Buffer API, which the Edge runtime doesn't
// provide - so we pin this route to the standard Node.js serverless runtime.
export const runtime = "nodejs";

// Every upload is different, so never let Vercel/Next cache this route.
export const dynamic = "force-dynamic";

export async function POST(request: NextRequest) {
  try {
    // The browser sends the file as multipart/form-data; formData() parses that for us
    const formData = await request.formData();
    const file = formData.get("file");

    // formData.get() can return a string (for text fields) or a File - make sure it's a File
    if (!file || typeof file === "string") {
      return NextResponse.json({ error: "No file was uploaded." }, { status: 400 });
    }

    const originalName = file.name; // e.g. "workshop_export.csv"
    const lowerName = originalName.toLowerCase();
    if (!lowerName.endsWith(".csv") && !lowerName.endsWith(".xlsx") && !lowerName.endsWith(".xls")) {
      return NextResponse.json({ error: "Only .csv, .xlsx or .xls files are supported." }, { status: 400 });
    }

    // Read the uploaded bytes into a Node Buffer, which is what the xlsx library expects
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    // Step 1: parse the raw file into an array of { column: value } row objects
    const inputRows = readInputFile(buffer, originalName);

    // Step 2: run the full cleaning pipeline (dates, phone numbers, refunds, etc.)
    const { rows, columnOrder } = processBda(inputRows);

    // Step 3: turn the cleaned rows back into an in-memory .xlsx file
    const outputBuffer = writeOutputWorkbook(rows, columnOrder);

    // Build "<original name>_cleaned.xlsx", same naming convention as the Python script
    const dotIndex = originalName.lastIndexOf(".");
    const baseName = dotIndex === -1 ? originalName : originalName.slice(0, dotIndex);
    const downloadName = `${baseName}_cleaned.xlsx`;

    // Send the workbook bytes straight back as a downloadable file
    return new NextResponse(outputBuffer, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${downloadName}"`,
      },
    });
  } catch (err: unknown) {
    // Surface a readable message to the frontend instead of a generic 500 page
    const message = err instanceof Error ? err.message : "Something went wrong while processing the file.";
    return NextResponse.json({ error: message }, { status: 400 });
  }
}
