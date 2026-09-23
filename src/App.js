import React, { useState } from 'react';
import logo from './logo.svg';
import './App.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig';

function App() {
  const { instance, accounts } = useMsal();
  const [apiData, setApiData] = useState(null);
  const [versionMostrada, setVersionMostrada] = useState("");

  const iniciarSesion = () => instance.loginRedirect(loginRequest).catch(console.error);
  const cerrarSesion = () => instance.logoutRedirect();

  // GET a la Versión 1 (Contrato original)
  const verCatalogoV1 = async () => {
    try {
      const responseAuth = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
      const responseApi = await fetch("https://y6omvgc6f0.execute-api.us-east-1.amazonaws.com/desarrollo/api/v1/productos", {
        method: "GET",
        headers: { "Authorization": `Bearer ${responseAuth.accessToken}`, "Content-Type": "application/json" }
      });
      const data = await responseApi.json();
      setApiData(data);
      setVersionMostrada("Versión 1 (Original)");
    } catch (error) { console.error("Error V1:", error); }
  };

  // GET a la Versión 2 (Contrato extendido con IVA)
  const verCatalogoV2 = async () => {
    try {
      const responseAuth = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
      const responseApi = await fetch("https://y6omvgc6f0.execute-api.us-east-1.amazonaws.com/desarrollo/api/v2/productos", {
        method: "GET",
        headers: { "Authorization": `Bearer ${responseAuth.accessToken}`, "Content-Type": "application/json" }
      });
      const data = await responseApi.json();
      setApiData(data);
      setVersionMostrada("Versión 2 (Con cálculo de IVA)");
    } catch (error) { console.error("Error V2:", error); }
  };

  // POST a la Versión 1
  const crearProductoPrueba = async () => {
    try {
      const responseAuth = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
      const nuevoProducto = { nombre: "Mouse Logitech G203", precio: 25000 };
      
      const responseApi = await fetch("https://y6omvgc6f0.execute-api.us-east-1.amazonaws.com/desarrollo/api/v1/productos", {
        method: "POST",
        headers: { "Authorization": `Bearer ${responseAuth.accessToken}`, "Content-Type": "application/json" },
        body: JSON.stringify(nuevoProducto)
      });

      if (responseApi.ok) {
        alert("¡Registro creado exitosamente en AWS RDS (V1)!");
        verCatalogoV1(); // Refresca la lista automáticamente
      } else {
        alert("Error al crear el registro. Revisa la consola.");
      }
    } catch (error) { console.error("Error POST:", error); }
  };

  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>Login en la plataforma con Azure Entra ID</p>

        <UnauthenticatedTemplate>
          <p>El usuario no está autenticado.</p>
          <button className="btn btn-success" onClick={iniciarSesion}>Iniciar sesión</button>
        </UnauthenticatedTemplate>

        <AuthenticatedTemplate>
          <h2>Usuario autenticado</h2>
          {accounts.length > 0 && (
            <>
              <p>Nombre: {accounts[0].name}</p>
              <p>Usuario: {accounts[0].username}</p>
            </>
          )}

          <div style={{ display: 'flex', gap: '15px', marginTop: '15px', justifyContent: 'center', flexWrap: 'wrap' }}>
            <button className="btn btn-primary" onClick={verCatalogoV1}>☁️ Ver Catálogo (V1)</button>
            <button className="btn btn-info" onClick={verCatalogoV2} style={{ backgroundColor: '#17a2b8', color: 'white', border: 'none' }}>🚀 Ver Catálogo (V2)</button>
            <button className="btn btn-warning" onClick={crearProductoPrueba} style={{ backgroundColor: '#ffc107', color: 'black', border: 'none' }}>➕ Crear Producto (V1)</button>
          </div>

          {apiData && (
            <div style={{ backgroundColor: "#282c34", padding: "15px", margin: "20px auto", borderRadius: "8px", border: "1px solid #61dafb", width: "80%", overflowX: "auto" }}>
              <h4 style={{ margin: "0 0 10px 0", color: "#61dafb" }}>Respuesta del Microservicio - {versionMostrada}:</h4>
              <pre style={{ textAlign: "left", fontSize: "14px", margin: 0 }}>
                {JSON.stringify(apiData, null, 2)}
              </pre>
            </div>
          )}

          <button className="btn btn-danger" onClick={cerrarSesion} style={{ margin: "15px" }}>Cerrar sesión</button>
        </AuthenticatedTemplate>
      </header>
    </div>
  );
}

export default App;