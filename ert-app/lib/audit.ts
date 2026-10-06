import type { Prisma } from "@prisma/client";
import { prisma } from "./db";

/** Never put passwords, tokens or phone numbers in `detail`. Use ids and counts. */
export async function audit(action: string, o: { userId?: string | null; email?: string | null; ip?: string; detail?: Record<string, unknown> } = {}) {
  try {
    await prisma.auditLog.create({ data: { action, userId: o.userId ?? null, email: o.email ?? null, ip: o.ip ?? null, detail: (o.detail ?? {}) as Prisma.InputJsonValue } });
  } catch (e) { console.error("audit write failed", e); }
}

export const recentCount = (userId: string, action: string, minutes: number) =>
  prisma.auditLog.count({ where: { userId, action, at: { gte: new Date(Date.now() - minutes * 60_000) } } });
