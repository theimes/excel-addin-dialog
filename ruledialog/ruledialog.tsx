import * as React from "react";

const Ruledialog: React.FC = () => {

  return (
    <div >

        <h1>Rule Dialog</h1>
        <p>This is a dialog for managing rules.</p>
        {/* Add your dialog content here */}
        <button onClick={() => Office.context.ui.messageParent("Hello from the dialog!")}>
            Send Message to Parent
        </button>  

    </div>
  );
};

export default Ruledialog;
