import { NextRequest, NextResponse } from "next/server";
import { fillModelErrorDetails, fillModelWorkbook } from "../../../server/fill-model/fill-model-service";

export const runtime = "nodejs";
export const maxDuration = 900;

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const query = String(formData.get("ticker") ?? "").trim();
    const file = formData.get("file");

    if (!query) return jsonError("Enter a ticker or company name.", 400);
    if (!(file instanceof File)) return jsonError("Upload an .xlsx or .xlsm workbook.", 400);

    const result = await fillModelWorkbook({
      query,
      workbookBuffer: await file.arrayBuffer()
    });
    const responseBody = result.output.buffer.slice(result.output.byteOffset, result.output.byteOffset + result.output.byteLength) as ArrayBuffer;

    return new NextResponse(responseBody, {
      headers: {
        "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "content-disposition": `attachment; filename="${result.outputName}"`,
        "x-output-filename": result.outputName,
        "x-fill-summary": encodeURIComponent(JSON.stringify(result.summary))
      }
    });
  } catch (error) {
    console.error(error);
    const { message, status } = fillModelErrorDetails(error);
    return jsonError(message, status);
  }
}

function jsonError(error: string, status: number) {
  return NextResponse.json({ error }, { status });
}
