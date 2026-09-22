import logo from './logo.svg';
import './App.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig';

function App() {

  const { instance, accounts } = useMsal();

  const iniciarSesion = () => {
    instance.loginRedirect(loginRequest)
      .catch(error => {
        console.error(error);
      });
  }

  const cerrarSesion = () => {
    instance.logoutRedirect();
  }


  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>
          Login en la plataforma con Azure Entra ID
        </p>


        <UnauthenticatedTemplate>
        <p>
          El usuario no está autenticado.
        </p>
        <button className="btn btn-success" onClick={iniciarSesion}>
        <i class="fa-solid fa-user"></i> Iniciar sesión
        </button>
        </UnauthenticatedTemplate>

        <AuthenticatedTemplate>
          <h2>Usuario autenticado</h2>
          {accounts.length > 0 && (
            <>
              <p>
                Nombre:
                {" "}
                {accounts[0].name}
              </p>
              <p>
                Usuario:
                {" "}
                {accounts[0].username}
              </p>
              <p>
                id:
                {" "}
                {accounts[0].idTokenClaims.oid}
              </p>
              <p>
                idTokenClaims:
                {" "}
                {JSON.stringify(accounts[0].idTokenClaims)}
              </p>
            </>
          )}
          <button className="btn btn-danger" onClick={cerrarSesion}>
          <i class="fa-solid fa-right-from-bracket"></i> Cerrar sesión
          </button>
        </AuthenticatedTemplate>


      </header>
    </div>
  );
}

export default App;
