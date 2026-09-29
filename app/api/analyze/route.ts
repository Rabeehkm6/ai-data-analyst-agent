export async function POST(req: Request) {
  const { rows } = await req.json();

  const totalSales = rows.reduce(
    (sum: number, row: any) =>
      sum + Number(row.Sales),
    0
  );

  return Response.json({
    insight: `
Dataset contains ${rows.length} records.
Total sales: AED ${totalSales}.
Top-performing regions should be prioritized.
Sales performance appears stable.
`,
  });
}