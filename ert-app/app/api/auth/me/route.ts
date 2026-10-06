import { NextResponse } from "next/server";
import { getSession } from "@/lib/auth";
export async function GET() {
  const s = await getSession();
  return s ? NextResponse.json({ email: s.email, role: s.role, siteId: s.siteId, mfaEnabled: s.mfaEnabled }) : NextResponse.json({ error: "Sign in required." }, { status: 401 });
}
