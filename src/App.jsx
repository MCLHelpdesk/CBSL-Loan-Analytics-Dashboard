import { useMsal, useIsAuthenticated } from '@azure/msal-react'
import { loginRequest } from './authConfig'
import Dashboard from './Dashboard'

function LoginPage() {
  const { instance } = useMsal()

  const handleLogin = () => {
    instance.loginRedirect(loginRequest)
  }

  return (
    <div className="login-page">

      <div className="background-shape orange-shape"></div>

      <div className="background-shape grey-shape"></div>

      <img
        src="/corner-block-logo.png"
        className="background-logo"
        alt=""
      />

      <div className="login-container">

        <img
          src="/corner-block-logo.png"
          className="main-logo"
          alt="Corner Block"
        />

        <div className="title-section">

          <h1>CBSL</h1>

          <h2>
            Loan Analytics Dashboard
          </h2>

        </div>

        <div className="login-card">

          <h3>
            Welcome
          </h3>

          <p>
            Sign in with your Microsoft 365 account
            <br />
            to continue.
          </p>

          <button
            className="microsoft-button"
            onClick={handleLogin}
          >

            <span className="microsoft-icon">

              <span></span>
              <span></span>
              <span></span>
              <span></span>

            </span>

            Sign in with Microsoft 365

          </button>

        </div>

        <div className="footer">

          © {new Date().getFullYear()}
          {' '}
          CBSL • Loan Analytics Dashboard

        </div>

      </div>

      <style>{`

        * {
          box-sizing: border-box;
        }

        body {
          margin: 0;
        }

        .login-page {

          min-height: 100vh;

          width: 100%;

          position: relative;

          overflow: hidden;

          display: flex;

          justify-content: center;

          align-items: center;

          font-family:
            Corbel,
            "Segoe UI",
            Arial,
            sans-serif;

          background:
            linear-gradient(
              135deg,
              #ffffff 0%,
              #f7f7f7 55%,
              #eeeeee 100%
            );

          color: #29323a;

        }

        .background-shape {

          position: absolute;

          z-index: 0;

          pointer-events: none;

        }

        .orange-shape {

          width: 600px;

          height: 600px;

          background: #f36c21;

          left: -400px;

          top: -230px;

          transform: rotate(45deg);

          opacity: 0.95;

        }

        .grey-shape {

          width: 700px;

          height: 90px;

          background: #55575a;

          right: -250px;

          bottom: 40px;

          transform: rotate(-38deg);

        }

        .background-logo {

          position: absolute;

          width: 520px;

          left: -80px;

          top: 120px;

          opacity: 0.08;

          transform: rotate(-15deg);

          z-index: 1;

        }

        .login-container {

          position: relative;

          z-index: 5;

          width: 100%;

          max-width: 700px;

          display: flex;

          flex-direction: column;

          align-items: center;

          padding: 30px 20px;

        }

        .main-logo {

          width: 125px;

          margin-bottom: 12px;

        }

        .title-section {

          text-align: center;

          margin-bottom: 35px;

        }

        .title-section h1 {

          margin: 0;

          font-size: 58px;

          line-height: 1;

          font-weight: 700;

          color: #f36c21;

          letter-spacing: -2px;

        }

        .title-section h2 {

          margin: 8px 0 0;

          font-size: 38px;

          line-height: 1.1;

          font-weight: 500;

          color: #29323a;

        }

        .login-card {

          width: 100%;

          max-width: 500px;

          padding: 45px 50px;

          background: rgba(255,255,255,0.97);

          border-radius: 18px;

          text-align: center;

          box-shadow:
            0 15px 45px rgba(0,0,0,0.12);

        }

        .login-card h3 {

          margin: 0 0 15px;

          font-size: 36px;

          font-weight: 600;

        }

        .login-card p {

          margin: 0 0 30px;

          font-size: 20px;

          line-height: 1.4;

          color: #55575a;

        }

        .microsoft-button {

          width: 100%;

          min-height: 58px;

          border: none;

          border-radius: 7px;

          background: #f36c21;

          color: white;

          font-family: inherit;

          font-size: 18px;

          font-weight: 600;

          cursor: pointer;

          display: flex;

          align-items: center;

          justify-content: center;

          gap: 14px;

          transition: 0.2s;

        }

        .microsoft-button:hover {

          background: #df5e18;

          transform: translateY(-2px);

        }

        .microsoft-icon {

          width: 22px;

          height: 22px;

          display: grid;

          grid-template-columns: 1fr 1fr;

          grid-template-rows: 1fr 1fr;

          gap: 2px;

        }

        .microsoft-icon span {

          background: white;

        }

        .footer {

          margin-top: 25px;

          font-size: 14px;

          color: #777;

        }

      `}</style>

    </div>
  )
}


function App() {

  const isAuthenticated =
    useIsAuthenticated()

  if (isAuthenticated) {

    return <Dashboard />

  }

  return <LoginPage />
}


export default App