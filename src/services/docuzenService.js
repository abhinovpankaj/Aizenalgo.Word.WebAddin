const AUTHENTICATIONBASEURL = "http://demo.aizenalgo.com:9016/api/WordProc/WordProcAuthentication";
const  VERIFICATIONBASEURL = "http://demo.aizenalgo.com:9016/api/WordProc/WordProcSessionDetails";

const SubmitDocumentService = ({sessionId,docId,type,uploadFile,fileName}) => {
    const endpoint = `${VERIFICATIONBASEURL}?SessionId=${sessionId}&DocID=${docId}&Mode=${type}`;
    var dataArray = new FormData();
    dataArray.append("fileName", fileName);
    dataArray.append("file", uploadFile);


    fetch(endpoint, {
       method: 'POST',
       body: dataArray,       
       headers: {
          'Content-type': 'multipart/form-data',
       },
    })
       .then((res) => {
        console.log(res);
        res.json();
    })
       
       .catch((err) => {
          console.log(err.message);
       });
 };

 export default SubmitDocumentService;