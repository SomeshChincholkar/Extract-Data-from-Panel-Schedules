import { useState, useEffect } from "react";
import { useNavigate } from "react-router-dom";
import axios from "axios";

function Demo2() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const navigate = useNavigate();

  useEffect(() => {
    const fetchData = async () => {
      try {
        const response = await axios.get('/api/demo2');
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
      <h2 className="text-2xl font-bold mb-4">Demo 2 Component</h2>
      <p className="text-gray-600 mb-4">{data?.message}</p>
      
      <h3 className="font-semibold mb-2">Features:</h3>
      <ul className="space-y-2">
        {data?.features.map((feature) => (
          <li key={feature.id} className="p-3 bg-gray-50 rounded">
            <h4 className="font-medium text-gray-800">{feature.title}</h4>
            <p className="text-gray-600">{feature.description}</p>
          </li>
        ))}
      </ul>
      
      <button
        onClick={() => navigate('/')}
        className="mt-6 px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600 transition-colors"
      >
        Back to Demo 1
      </button>
      
      <p className="text-sm text-gray-500 mt-4">
        Timestamp: {new Date(data?.timestamp).toLocaleString()}
      </p>
    </div>
  );
}

export default Demo2; 