import { IPublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { Authenticate } from "./components/Authenticate";
import ProvideAppContext from "./context";

// <AppPropsSnippet>
type AppProps = {
  pca: IPublicClientApplication
};
// </AppPropsSnippet>

function App({ pca }: AppProps) {
  return (
    <MsalProvider instance={pca}>
      <ProvideAppContext>
        <Authenticate />
      </ProvideAppContext>
    </MsalProvider>
  );
}

export default App;
