const AUTHENTICATIONBASEURL = "https://demo.aizenalgo.com:9016/api/WordProc/WordProcAuthentication";
const VERIFICATIONBASEURL = "https://demo.aizenalgo.com:9016/api/WordProc/WordProcSessionDetails";

const DocuzenSessionVerification = ({stoken,dvid,uploadFileData,fileName},type)=>{
  const endpoint = `${VERIFICATIONBASEURL}?SessionId=${stoken}&DocID=${dvid}&Mode=${type}`;
  var file = new Blob([uploadFileData], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
  var dataArray = new FormData();
  dataArray.append("fileName", fileName);
  dataArray.append("file", file);
  
  fetch(endpoint, {
    method: 'POST',
    body: dataArray
  })
  .then(response => {
    if (!response.ok) throw (`invalid response: ${response.status}`);
        return response.json()
    })
  .then(data => {
    console.log(data);
    return data;
  })
  .catch((err) => {
      console.log(err);
      return err;
    });
}

const DocuzenAuthentication = ({stoken,dvid,uploadFileData,fileName},type,userName, password,)=>{
  
  const endpoint = `${AUTHENTICATIONBASEURL}?UserName=${userName}&Password=${password}&SessionId=${stoken}&DocID=${dvid}&Mode=${type}`;
  var dataArray = new FormData();
  var file = new Blob([uploadFileData], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
  dataArray.append("fileName", fileName);
  dataArray.append("file", file);
    
  fetch(endpoint, {
    method: 'POST',
    body: dataArray
  })
  .then(response => {
    if (!response.ok) throw (`invalid response: ${response.status}`);
        return response.json()
    }) 
  .then(data => {
    console.log(data);
    return data;
  })
  .catch((err) => {
      console.log(err);
      return err;
    });
}

export {DocuzenAuthentication,DocuzenSessionVerification};