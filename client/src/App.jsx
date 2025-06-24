import { useState } from "react";

export default function App() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleUpload = async () => {
    if (!file) return alert("Please select a file.");
    setLoading(true);

    const formData = new FormData();
    formData.append("image", file);

    try {
      const res = await fetch("http://localhost:5000/upload", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const errText = await res.text();
        throw new Error(errText);
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Panel_Board_Listing.xlsx";
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      alert("Upload failed: " + err.message);
    }
    setLoading(false);
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gray-100 p-4">
      <div className="bg-white shadow-xl rounded-2xl p-8 w-full max-w-md text-center">
        <h1 className="text-2xl font-bold text-blue-600 mb-4">Extract Data from Panel Schedules</h1>

        <input
          type="file"
          accept="image/*"
          onChange={(e) => setFile(e.target.files[0])}
          className="mb-4 w-full"
        />

        <button
          onClick={handleUpload}
          disabled={loading}
          className={`w-full py-2 px-4 text-white font-semibold rounded-xl ${
            loading ? "bg-blue-300 cursor-not-allowed" : "bg-blue-600 hover:bg-blue-700"
          }`}
        >
          {loading ? "Processing..." : "Upload & Generate Excel"}
        </button>
      </div>
    </div>
  );
}