import { useState, useEffect } from "react";
import { useNavigate } from "react-router-dom";
import axios from "axios";

function Demo() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const navigate = useNavigate();

  useEffect(() => {
    const fetchData = async () => {
      try {
        const response = await axios.get('/api/demo');
        setData(response.data);
        setLoading(false);
      } catch (err) {
        console.error('Error details:', err);
        setError(err.message);
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  if (loading) return <div className="text-center">Loading...</div>;
  if (error) return <div className="text-red-500">Error: {error}</div>;

  return (
    <div className="max-w-md mx-auto mt-10 p-6 bg-white rounded-lg shadow-lg">
      <h2 className="text-2xl font-bold mb-4">Demo Component</h2>
      <p className="text-gray-600 mb-4">{data?.message}</p>
      
      <h3 className="font-semibold mb-2">Items:</h3>
      <ul className="list-disc pl-5">
        {data?.items.map((item) => (
          <li key={item.id} className="text-gray-700">{item.name}</li>
        ))}
      </ul>
      
      <button
        onClick={() => navigate('/demo2')}
        className="mt-6 px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600 transition-colors"
      >
        Go to Demo 2
      </button>
      
      <p className="text-sm text-gray-500 mt-4">
        Timestamp: {new Date(data?.timestamp).toLocaleString()}
      </p>
    </div>
  );
}

export default Demo; 