// Run once: SEED_ADMIN_EMAIL=you@company.com SEED_ADMIN_PASSWORD='long-passphrase' npx tsx prisma/seed.ts
import { PrismaClient } from "@prisma/client";
import bcrypt from "bcryptjs";
const prisma = new PrismaClient();
async function main() {
  const email = process.env.SEED_ADMIN_EMAIL?.toLowerCase(), pw = process.env.SEED_ADMIN_PASSWORD;
  if (!email || !pw || pw.length < 12) throw new Error("Set SEED_ADMIN_EMAIL and SEED_ADMIN_PASSWORD (12+ chars).");
  await prisma.user.upsert({ where: { email }, update: {}, create: { email, passwordHash: await bcrypt.hash(pw, 12), role: "SUPER_ADMIN" } });
  console.log("Super admin ready:", email);
}
main().finally(() => prisma.$disconnect());
