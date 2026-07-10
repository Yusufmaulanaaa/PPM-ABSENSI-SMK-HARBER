import axios from 'axios';

const API_BASE_URL = import.meta.env.VITE_API_URL || 'https://backend-production-1924.up.railway.app/api';

const api = axios.create({
  baseURL: API_BASE_URL,
});

api.interceptors.request.use((config) => {
  const token = localStorage.getItem('absensi_token');
  if (token) {
    config.headers.Authorization = `Bearer ${token}`;
  }
  return config;
});

api.interceptors.response.use(
  (response) => response,
  (error) => {
    if (error.response?.status === 401) {
      localStorage.removeItem('absensi_token');
      localStorage.removeItem('absensi_user');
      if (!window.location.pathname.includes('/login')) {
        window.location.href = '/login';
      }
    }
    return Promise.reject(error);
  }
);

export const UPLOADS_BASE_URL = API_BASE_URL.replace(/\/api$/, '');

export default api;
