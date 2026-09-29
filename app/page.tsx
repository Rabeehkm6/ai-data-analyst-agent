import CsvUploader from "@/components/CsvUploader";

export default function Home() {
  return (
    <main className="min-h-screen bg-slate-100 p-10">
  <div className="max-w-[1400px] mx-auto">
    <h1 className="text-5xl font-bold mb-8">
      AI Data Analyst Agent
    </h1>

    <CsvUploader />
  </div>
</main>
  );
}