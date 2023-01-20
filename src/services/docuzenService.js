const AUTHENTICATIONBASEURL = "https://demo.aizenalgo.com:9016/api/WordProc/WordProcAuthentication";
const VERIFICATIONBASEURL = "https://demo.aizenalgo.com:9016/api/WordProc/WordProcSessionDetails";

const SubmitDocumentService = ({stoken,dvid,uploadFile,fileName},type) => {
    const endpoint = `${VERIFICATIONBASEURL}?SessionId=${stoken}&DocID=${dvid}&Mode=${type}`;
    var dataArray = new FormData();
    dataArray.append("fileName", fileName);
    dataArray.append("file", uploadFile);


    fetch(endpoint, {
       method: 'POST',
       body: dataArray,       
       headers: {
         'Access-Control-Allow-Origin': '*',          
       },
    })
       .then(response => {
         if (!response.ok) throw (`invalid response: ${response.status}`); 
         return response.json()
    })
    .then(data => console.log(data))
       
       .catch((err) => {
          console.log(err);
       });
 };

 export default SubmitDocumentService;