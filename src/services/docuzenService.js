const AUTHENTICATIONBASEURL = "https://demo.aizenalgo.com:9016/api/WordProc/WordProcAuthentication";
const VERIFICATIONBASEURL = "https://demo.aizenalgo.com:9016/api/WordProc/WordProcSessionDetails";

const DocuzenSessionVerification = async ({stoken,dvid,uploadFileData,fileName},type)=>{
  const endpoint = `${VERIFICATIONBASEURL}?SessionId=${stoken}&DocID=${dvid}&Mode=${type}`;
  var file = new Blob([uploadFileData], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
  var dataArray = new FormData();
  dataArray.append("fileName", fileName);
  dataArray.append("file", file);
  
  try {
    const response = await fetch(endpoint, {
      method: 'POST',
      body: dataArray
    });
    if (!response.ok)
      throw (`invalid response: ${response.status}`);
    const data = await response.json();
    console.log(data);
    return await Promise.resolve(data);
  } catch (err) {
    console.log(err);
    return await Promise.reject(err);
  }
}

const DocuzenAuthentication = async ({stoken,dvid,uploadFileData,fileName},type,userName, password,)=>{
  
  const endpoint = `${AUTHENTICATIONBASEURL}?UserName=${userName}&Password=${password}&SessionId=${stoken}&DocID=${dvid}&Mode=${type}`;
  var dataArray = new FormData();
  var file = new Blob([uploadFileData], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
  dataArray.append("fileName", fileName);
  dataArray.append("file", file);
    
  try {
    const response = await fetch(endpoint, {
      method: 'POST',
      body: dataArray
    });
    if (!response.ok)
      throw (`invalid response: ${response.status}`);
    const data = await response.json();
    console.log(data);
    return data;
  } catch (err) {
    console.log(err);
    return err;
  }
}

export {DocuzenAuthentication,DocuzenSessionVerification};