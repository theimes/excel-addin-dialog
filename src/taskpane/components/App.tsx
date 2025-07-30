import * as React from "react";
import Header from "./Header";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { insertText } from "../taskpane";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Hi" />
      {/* 
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
       */}
      <TextInsertion insertText={insertText} />
      {/* Open dialog button */}
      <button onClick={() => Office.context.ui.displayDialogAsync(
        "https://localhost:3000/rules/", 
        { 
          height: 50, 
          width: 50
        },
        (asyncResult: Office.AsyncResult<Office.Dialog>) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to open dialog: " + asyncResult.error.message);
            return;
          }
          const dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
            console.log("Message from dialog: " + (JSON.stringify(args)));
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (args) => {
            console.log("Dialog event received: " + (JSON.stringify(args)));
          });
        }
      )
        }>
        Open Dialog
      </button>
    </div>
  );
};

export default App;
