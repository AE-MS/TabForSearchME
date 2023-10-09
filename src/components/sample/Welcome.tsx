import { useContext, useState } from "react";
import {
  Image,
  SelectTabEvent,
  SelectTabData,
  TabValue,
} from "@fluentui/react-components";
import "./Welcome.css";
import { useData } from "@microsoft/teamsfx-react";
import { TeamsFxContext } from "../Context";
import { FrameContexts, app, authentication, dialog, version } from "@microsoft/teams-js";

function submitAndRequestUrlDialog() {
  dialog.url.submit({ data: "requestUrl" });
}

function submitAndRequestCardDialog() {
  dialog.url.submit({ data: "requestCard" });
}

function submitAndRequestMessageDialog() {
  dialog.url.submit({ data: "requestMessage" });
}

function submitAndRequestNoResponse() {
  dialog.url.submit({ data: "requestNoResponse" });
}

function submitConfig() {
  authentication.notifySuccess((document.getElementById('configValue') as HTMLInputElement).value)
}

export function Welcome(props: { showFunction?: boolean; environment?: string }) {
  const { showFunction, environment } = {
    showFunction: true,
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";

  const [selectedValue, setSelectedValue] = useState<TabValue>("local");

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedValue(data.value);
  };
  const { teamsUserCredential } = useContext(TeamsFxContext);
  const { loading, data, error } = useData(async () => {
    if (teamsUserCredential) {
      const userInfo = await teamsUserCredential.getUserInfo();
      return userInfo;
    }
  });
  const userName = loading || error ? "" : data!.displayName;
  const context = useData(async () => {
    await app.initialize();
    const context = await app.getContext();
    return context;
  })?.data;

  const hubName: string | undefined = context?.app.host.name;
  const frameContext: FrameContexts | undefined = context?.page.frameContext;
  const cardDialogsIsSupported: boolean | undefined = context === undefined ? undefined : dialog.adaptiveCard.isSupported();

  const urlParams = new URLSearchParams(window.location.search);
  const pageQueryParamValue = urlParams.get('page');

  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="hello.png" />
        <h1 className="center">Congratulations{userName ? ", " + userName : ""}!</h1>
        {hubName && <p className="center">This is a URL dialog running in {hubName}</p>}
        <p className="center">Your app is running in your {friendlyEnvironmentName}</p>
        <p className="center">TeamsJS version: {version}</p>
        <p className="center">Card Dialogs is supported: {cardDialogsIsSupported ? "true" : "false"}</p>
        <p className="center">The context frame context is {frameContext}</p>
        <p className="center">The current URL is {window.location.href}</p>

        { frameContext === FrameContexts.task && (
          <div>
            <button onClick={submitAndRequestUrlDialog}>Submit And Request URL Dialog</button>
            <button onClick={submitAndRequestCardDialog}>Submit And Request Card Dialog</button>
            <button onClick={submitAndRequestMessageDialog}>Submit And Request Message Dialog</button>
            <button onClick={submitAndRequestNoResponse}>Submit And Request No Response</button>
          </div>
        )}

        { pageQueryParamValue === 'config' && (
          <div>
            <input id="configValue" type="text" placeholder="Enter your config value" />
            <input type="button" value="Submit Configuration" onClick={submitConfig} />
          </div>
        )}

      </div>
    </div>
  );
}
