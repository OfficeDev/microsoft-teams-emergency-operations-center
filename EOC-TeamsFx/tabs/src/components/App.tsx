import { Loader, Provider, ThemeInput, teamsTheme } from "@fluentui/react-northstar";
import { useTeamsUserCredential } from "@microsoft/teamsfx-react";
import { Redirect, Route, HashRouter as Router } from "react-router-dom";
import { TeamsFxContext } from "./Context";
import Tab from "./Tab";
import TabConfig from './TabConfig';

const startLoginPageUrl = process.env.REACT_APP_START_LOGIN_PAGE_URL;
const clientId = process.env.REACT_APP_CLIENT_ID;
/**
 * The main app which handles the initialization and routing
 * of the app.
 */
export default function App() {
  const { loading, theme, themeString, teamsUserCredential } = useTeamsUserCredential({
    initiateLoginEndpoint: startLoginPageUrl!,
    clientId: clientId!,
  });
  return (
    <TeamsFxContext.Provider value={{ theme, themeString, teamsUserCredential }}>
      <Provider theme={(theme as ThemeInput) || teamsTheme} styles={{ backgroundColor: "#eeeeee" }}>
        <Router>
          <Route exact path="/">
            <Redirect to="/tab" />
          </Route>
          {loading ? (
            <Loader style={{ margin: 100 }} />
          ) : (
            <>
              <Route exact path="/tab" component={Tab} />
              <Route exact path="/config" component={TabConfig} />
            </>
          )}
        </Router>
      </Provider>
    </TeamsFxContext.Provider>
  );
}
