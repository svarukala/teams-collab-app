import * as React from 'react';
import { useState, useEffect } from 'react';
import { List } from "@fluentui/react-northstar";

function SPOReusable(props) {
  // Declare a new state variable, which we'll call "count"
  const [ssoToken, setSsoToken] = useState<string>();
  const [error, setError] = useState<string>();
  const [resourceOboToken, setResourceOboToken] = useState<string>();
  const [searchResults, setSearchResults] = useState<any[]>();

  useEffect(() => {
    if (props.idToken) {
      setSsoToken(props.idToken);
    }
  }, []);
  
  const getResourceAccessTokenOBO = async () => {
    fetch(`https://azfun.ngrok.io/api/TeamsOBOHelper?ssoToken=${ssoToken}&tokenFor=spo`).then(async response =>{
        const responsePayload = await response.json();
        if (response.ok) {
            setResourceOboToken(responsePayload.access_token);
        } else {
          if (responsePayload!.error === "consent_required") {
            setError("consent_required");
          } else {
            setError("unknown SSO error");
          }
        }
    })
    .catch((err) => {
        console.log(err);
        console.log("OBO Failed");
        setError("OBO Failed: " + err);
    });
  };

  useEffect(() => {
    if (ssoToken && ssoToken.length > 0) {
      getResourceAccessTokenOBO();
    }
  }, [ssoToken]);      

  useEffect(() => {
    getSPOSearchResutls();
  }, [resourceOboToken]);

    const getSPOSearchResutls = async () => {
        if (!resourceOboToken) { return; }
    
        const endpoint = `https://m365x229910.sharepoint.com/_api/search/query?querytext=%27*%27&selectproperties=%27Author,Path,Title,Url%27&rowlimit=10`;
        const requestObject = {
        method: 'GET',
        headers: {
            "authorization": "bearer " + resourceOboToken,
            "accept": "application/json; odata=nometadata"
        }
        };
    
        const response = await fetch(endpoint, requestObject);
        const responsePayload = await response.json();
    
        console.log(responsePayload.value);
        if (response.ok) {
            const resultSet = responsePayload.PrimaryQueryResult.RelevantResults.Table.Rows.map((result: any) => ({
            key:result.Cells[8].Value,
            //header:result.Cells[0].Value,
            headerMedia:result.Cells[2].Value,
            content:result.Cells[1].Value,
            }));        
        console.log(JSON.stringify(resultSet));
        setSearchResults(resultSet);
    }
    }

  return (
    <div>
        {error && "Error: " + error}
        <br />
        {searchResults && <List items={searchResults} />}
        <br/>
        {ssoToken && "ID Token: " + ssoToken}
        <br/><br/>
        {resourceOboToken && "Access Token: " + resourceOboToken}
    </div>
  );
}

export default SPOReusable;


/*
<p>You clicked {count} times</p>
<button onClick={() => setCount(count + 1)}>
  Click me
</button>
*/