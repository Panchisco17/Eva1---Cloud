import React, { useState } from 'react';
import './App.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig';

function App() {
  const { instance, accounts } = useMsal();
  const [apiData, setApiData] = useState(null);
  const [versionMostrada, setVersionMostrada] = useState("");
  const [vistaActiva, setVistaActiva] = useState(null);

  const [tituloSolicitud, setTituloSolicitud] = useState("");
  const [descripcionSolicitud, setDescripcionSolicitud] = useState("");
  const [categoriaSolicitud, setCategoriaSolicitud] = useState("Hardware");
  const [prioridadSolicitud, setPrioridadSolicitud] = useState("Media");

  const [nombreProducto, setNombreProducto] = useState("");
  const [precioProducto, setPrecioProducto] = useState("");

  const userRoles = accounts[0]?.idTokenClaims?.roles || [];
  const esCliente = userRoles.includes("ROLE_CLIENTE");
  const esOperador = userRoles.includes("ROLE_OPERADOR");
  const esAdmin = userRoles.includes("ROLE_ADMINISTRADOR");
  const puedeGestionar = esOperador || esAdmin;

  const iniciarSesion = () => instance.loginRedirect(loginRequest).catch(console.error);
  const cerrarSesion = () => instance.logoutRedirect();

  // Función centralizada para optimizar y reducir los fetch repetitivos
  const realizarPeticion = async (endpoint, method = "GET", body = null, isJson = true) => {
    try {
      const responseAuth = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
      const url = `https://y6omvgc6f0.execute-api.us-east-1.amazonaws.com/desarrollo/api/${endpoint}`;
      
      const options = {
        method,
        headers: { "Authorization": `Bearer ${responseAuth.accessToken}`, "Content-Type": "application/json" }
      };

      if (body) options.body = isJson ? JSON.stringify(body) : body;

      const response = await fetch(url, options);
      if (!response.ok) throw new Error("Error en la operación");
      
      const text = await response.text();
      return text ? JSON.parse(text) : true;
    } catch (error) {
      console.error(`Error en ${method} ${endpoint}:`, error);
      return null;
    }
  };

  const verCatalogo = async (version) => {
    setVistaActiva(null);
    const data = await realizarPeticion(`${version}/productos`);
    if (data) {
      setApiData(data);
      setVersionMostrada(version === "v1" ? "Catálogo de Productos" : "Catálogo Valorizado (IVA)");
    }
  };

  const guardarProducto = async (e) => {
    e.preventDefault();
    if (!nombreProducto.trim() || !precioProducto) return alert("Completa todos los campos.");
    
    const exito = await realizarPeticion("v1/productos", "POST", { nombre: nombreProducto, precio: parseFloat(precioProducto) });
    if (exito) {
      alert("Producto registrado exitosamente.");
      setNombreProducto(""); setPrecioProducto(""); verCatalogo("v1");
    } else alert("Error al registrar el producto.");
  };

  const verSolicitudes = async () => {
    setVistaActiva(null);
    const endpoint = esCliente ? `v1/solicitudes/usuario/${accounts[0].username}` : `v1/solicitudes`;
    const data = await realizarPeticion(endpoint);
    if (data) {
      setApiData(data);
      setVersionMostrada(esCliente ? "Mis Solicitudes" : "Gestión de Solicitudes");
    }
  };

  const guardarSolicitud = async (e) => {
    e.preventDefault();
    if (!tituloSolicitud.trim() || !descripcionSolicitud.trim()) return alert("Completa título y descripción.");
    
    const nuevaSolicitud = {
      titulo: tituloSolicitud, descripcion: descripcionSolicitud,
      categoria: categoriaSolicitud, prioridad: prioridadSolicitud,
      usuarioSolicitante: accounts[0].username, estado: "CREADA"
    };

    const exito = await realizarPeticion("v1/solicitudes", "POST", nuevaSolicitud);
    if (exito) {
      alert("Solicitud creada con éxito.");
      setTituloSolicitud(""); setDescripcionSolicitud(""); verSolicitudes();
    } else alert("Error al crear la solicitud.");
  };

  // Máquina de estados unificada
  const avanzarEstadoSolicitud = async (id, estadoActual) => {
    let nuevoEstado = "";
    let versionApi = "v1";

    switch (estadoActual?.toUpperCase()) {
      case "CREADA": nuevoEstado = "ASIGNADA"; break;
      case "ASIGNADA": nuevoEstado = "EN_PROCESO"; break;
      case "EN_PROCESO": 
        nuevoEstado = "RESUELTA"; 
        versionApi = "v2"; 
        break;
      case "RESUELTA": nuevoEstado = "CERRADA"; break;
      default: return;
    }

    const body = versionApi === "v1" ? nuevoEstado : { estado: nuevoEstado };
    const isJson = versionApi === "v2";
    
    const exito = await realizarPeticion(`${versionApi}/solicitudes/${id}/estado`, "PUT", body, isJson);
    if (exito) verSolicitudes();
    else alert("El backend rechazó el cambio de estado.");
  };

  const eliminarSolicitud = async (id) => {
    if (!window.confirm("¿Seguro que deseas eliminar esta solicitud?")) return;
    const exito = await realizarPeticion(`v1/solicitudes/${id}`, "DELETE");
    if (exito) verSolicitudes();
    else alert("Error de permisos. Acción reservada para Administradores.");
  };

  const renderizarTabla = () => {
    if (!apiData) return null;
    if (!Array.isArray(apiData)) return <div className="detalle-registro"><pre>{JSON.stringify(apiData, null, 2)}</pre></div>;
    if (apiData.length === 0) return <p className="empty-msg">No hay registros disponibles en este momento.</p>;

    const esProducto = versionMostrada.includes("Catálogo");

    if (esProducto) {
      const esV2 = apiData[0].hasOwnProperty('precio_iva');
      return (
        <table className="modern-table">
          <thead>
            <tr>
              <th>ID</th>
              <th>Nombre del Producto</th>
              <th>Valor Unitario</th>
              {esV2 && <th>Valor Total (IVA inc.)</th>}
            </tr>
          </thead>
          <tbody>
            {apiData.map(item => {
              const precioBase = item.precio !== undefined ? item.precio : item.precio_neto;
              return (
                <tr key={item.id}>
                  <td>{item.id}</td>
                  <td>{item.nombre}</td>
                  <td>{precioBase != null ? `$${precioBase.toLocaleString()}` : 'N/A'}</td>
                  {esV2 && <td>{item.precio_iva != null ? `$${item.precio_iva.toLocaleString()} ${item.moneda || ''}` : 'N/A'}</td>}
                </tr>
              );
            })}
          </tbody>
        </table>
      );
    } else {
      return (
        <table className="modern-table">
          <thead>
            <tr>
              <th>Folio</th>
              <th>Título</th>
              <th>Categoría</th>
              <th>Prioridad</th>
              <th>Estado</th>
              {!esCliente && <th>Solicitante</th>}
              <th>Fecha Registro</th>
              {puedeGestionar && <th>Acción</th>}
            </tr>
          </thead>
          <tbody>
            {apiData.map(item => {
              const estado = item.estado?.toUpperCase();
              return (
                <tr key={item.id}>
                  <td>#{item.id}</td>
                  <td style={{ fontWeight: '500' }}>{item.titulo}</td>
                  <td>{item.categoria}</td>
                  <td>{item.prioridad}</td>
                  <td><span className={`status-badge status-${estado?.toLowerCase()}`}>{estado}</span></td>
                  {!esCliente && <td>{item.usuarioSolicitante}</td>}
                  <td>{item.fechaCreacion ? new Date(item.fechaCreacion).toLocaleString() : 'N/A'}</td>
                  {puedeGestionar && (
                    <td>
                      {estado === 'CREADA' && <button className="btn-action btn-asignar" onClick={() => avanzarEstadoSolicitud(item.id, estado)}>Asignar</button>}
                      {estado === 'ASIGNADA' && <button className="btn-action btn-proceso" onClick={() => avanzarEstadoSolicitud(item.id, estado)}>Iniciar Trabajo</button>}
                      {estado === 'EN_PROCESO' && <button className="btn-action btn-resolver" onClick={() => avanzarEstadoSolicitud(item.id, estado)}>Resolver</button>}
                      {estado === 'RESUELTA' && <button className="btn-action btn-cerrar" onClick={() => avanzarEstadoSolicitud(item.id, estado)}>Cerrar</button>}
                      {estado === 'CERRADA' && <span style={{ fontSize: '0.8rem', color: '#94a3b8' }}>Finalizado</span>}
                      {esAdmin && <button className="btn-action btn-eliminar" style={{marginLeft: '5px'}} onClick={() => eliminarSolicitud(item.id)}>🗑️</button>}
                    </td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      );
    }
  };

  return (
    <>
      <UnauthenticatedTemplate>
        <div className="login-container">
          <div className="login-card">
            <h1>MesaTech <span>Cloud</span></h1>
            <tr>  </tr>
            <button className="btn-login" onClick={iniciarSesion}>Ingresar</button>
          </div>
        </div>
      </UnauthenticatedTemplate>

      <AuthenticatedTemplate>
        <div className="dashboard-layout">
          <aside className="sidebar">
            <div className="sidebar-brand"><h2>MesaTech <span>Cloud</span></h2></div>
            {accounts.length > 0 && (
              <div className="user-profile">
                <span className="user-name">{accounts[0].name}</span>
                <span className="user-email" style={{ fontSize: '0.75rem', color: '#94a3b8', display: 'block', marginTop: '4px' }}>{accounts[0].username}</span>
                <span className="user-role">{esAdmin ? "Administrador" : esOperador ? "Operador" : "Cliente"}</span>
              </div>
            )}
            <nav className="sidebar-nav">
              <div className="nav-section">
                <h3>Mesa de Ayuda</h3>
                <button className="btn-nav primary" onClick={() => { setVistaActiva('formSolicitud'); setApiData(null); }}>Crear Solicitud</button>
                <button className="btn-nav" onClick={verSolicitudes}>{esCliente ? "Mis Solicitudes" : "Gestión de Solicitudes"}</button>
              </div>
              <div className="nav-section">
                <h3>Inventario</h3>
                <button className="btn-nav" onClick={() => verCatalogo("v1")}>Catálogo de Productos</button>
                <button className="btn-nav" onClick={() => verCatalogo("v2")}>Catálogo Valorizado</button>
                {esAdmin && <button className="btn-nav primary" onClick={() => { setVistaActiva('formProducto'); setApiData(null); }} style={{ marginTop: '5px' }}>Registrar Producto</button>}
              </div>
            </nav>
            <div className="sidebar-footer"><button className="btn-logout" onClick={cerrarSesion}>Cerrar Sesión</button></div>
          </aside>

          <main className="main-content">
            <header className="topbar"><h2>Panel de Control</h2></header>
            <div className="content-area">
              {vistaActiva === 'formSolicitud' && (
                <div className="form-card">
                  <h3>Generar Ticket de Soporte</h3>
                  <form onSubmit={guardarSolicitud}>
                    <div className="form-group">
                      <label>Asunto del Ticket:</label>
                      <input type="text" className="form-input" value={tituloSolicitud} onChange={(e) => setTituloSolicitud(e.target.value)} required />
                    </div>
                    <div className="form-group">
                      <label>Detalle Técnico:</label>
                      <textarea className="form-input" rows="3" value={descripcionSolicitud} onChange={(e) => setDescripcionSolicitud(e.target.value)} required></textarea>
                    </div>
                    <div className="form-group">
                      <label>Categoría:</label>
                      <select className="form-select" value={categoriaSolicitud} onChange={(e) => setCategoriaSolicitud(e.target.value)}>
                        <option value="Hardware">Hardware</option><option value="Software">Software</option><option value="Redes">Redes</option>
                      </select>
                    </div>
                    <div className="form-group">
                      <label>Nivel de Prioridad:</label>
                      <select className="form-select" value={prioridadSolicitud} onChange={(e) => setPrioridadSolicitud(e.target.value)}>
                        <option value="Baja">Baja</option><option value="Media">Media</option><option value="Alta">Alta</option>
                      </select>
                    </div>
                    <div className="form-actions">
                      <button type="button" className="btn-cancel" onClick={() => setVistaActiva(null)}>Cancelar</button>
                      <button type="submit" className="btn-submit">Procesar Solicitud</button>
                    </div>
                  </form>
                </div>
              )}

              {vistaActiva === 'formProducto' && esAdmin && (
                <div className="form-card">
                  <h3>Alta de Nuevo Producto</h3>
                  <form onSubmit={guardarProducto}>
                    <div className="form-group">
                      <label>Descripción del Ítem:</label>
                      <input type="text" className="form-input" value={nombreProducto} onChange={(e) => setNombreProducto(e.target.value)} required />
                    </div>
                    <div className="form-group">
                      <label>Valor Unitario (CLP):</label>
                      <input type="number" className="form-input" value={precioProducto} onChange={(e) => setPrecioProducto(e.target.value)} required />
                    </div>
                    <div className="form-actions">
                      <button type="button" className="btn-cancel" onClick={() => setVistaActiva(null)}>Cancelar</button>
                      <button type="submit" className="btn-submit">Registrar en Catálogo</button>
                    </div>
                  </form>
                </div>
              )}

              {!apiData && !vistaActiva && (
                <div className="welcome-msg">
                  <h3>Visión General</h3><p>Seleccione una operación en el panel lateral.</p>
                </div>
              )}

              {apiData && !vistaActiva && (
                <div className="response-card table-container">
                  <h4 style={{ borderBottom: '2px solid #e2e8f0', paddingBottom: '10px', marginBottom: '20px', color: '#1e293b' }}>{versionMostrada}</h4>
                  {renderizarTabla()}
                </div>
              )}
            </div>
          </main>
        </div>
      </AuthenticatedTemplate>
    </>
  );
}

export default App;