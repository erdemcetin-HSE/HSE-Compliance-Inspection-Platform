import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function main() {
  await prisma.ptwAttachment.deleteMany();
  await prisma.ptwTeamChange.deleteMany();
  await prisma.ptwDailyLog.deleteMany();
  await prisma.ptwPrecaution.deleteMany();
  await prisma.ptwPreWorkCheck.deleteMany();
  await prisma.ptwTeamMember.deleteMany();
  await prisma.ptwEquipment.deleteMany();
  await prisma.ptwSafetySystem.deleteMany();
  await prisma.ptwHazard.deleteMany();
  await prisma.training.deleteMany();
  await prisma.riskRecord.deleteMany();
  await prisma.observation.deleteMany();
  await prisma.ptwRecord.deleteMany();
  await prisma.ppeTransaction.deleteMany();
  await prisma.ppeItem.deleteMany();
  await prisma.inspectionAttachment.deleteMany();
  await prisma.inspection.deleteMany();
  await prisma.employee.deleteMany();

  console.log('Seed completed: demo data cleared.');
}

main()
  .catch((error) => {
    console.error(error);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
