import * as React from 'react';
import * as ReactDOM from "react-dom";

import 'antd/dist/antd.css';
import styles from './UploadFile.module.scss';
//import { escape } from '@microsoft/sp-lodash-subset';
import { Dropzone, FileItem } from "@dropzone-ui/react"; //https://dropzone-ui.herokuapp.com/
import { useState } from 'react';
import { DefaultButton, IIconProps } from '@fluentui/react';
import { IUploadFileProps } from './IUploadFile';
import { Progress, Card } from 'antd';
import { Icon } from '@fluentui/react/lib/Icon';
import swal from 'sweetalert2';
import { LogHelper } from '../../helpers/LogHelper';

const OneDriveIcon = () => <Icon style={{ color: '#006ac0', fontSize: '30px' }} iconName="OneDriveFolder16" />;

const volume0Icon: IIconProps = {
  iconName: 'CloudUpload',
  styles: {
    root: {
      fontSize: '16px'
    }
  }

};


export const UploadFile: React.FC<IUploadFileProps> = (props) => {


  const [files, setFiles] = useState([]);
  const [isUploadFinished, setIsUploadFinished] = useState(false);
  const [uploadedFile, setUploadedFile] = useState<any>({});
  const [showProgressBar, setShowProgressBar] = useState<any>(false);
  const [percentComplete, setPercentComplete] = React.useState(0);
  const [isUploading, setIsUploading] = useState(false);



  const updateFiles = (incommingFiles): void => {
    //do something with the files
    setFiles(incommingFiles);
    //even your own upload implementation
  };
  const handleDelete = (id): void => {
    setFiles(files.filter((x) => x.id !== id));
  };

  const delay = time => new Promise(res => setTimeout(res, time));

  const showSucessAndReload = (messageToShow: string): void => {
    swal({
      title: 'All done.',
      text: messageToShow,
      type: 'success',
      confirmButtonColor: "#3085d6"
    }).then(() => {
      location.reload();
    });

  }

  const handleUploadFile = async () => {
    //await delay(5000);
    setShowProgressBar(true);
    let result: any = null
    const file: File = files[0].file;

    try {

      //payload for OneDrive
      const payload = {
        "@microsoft.graph.conflictBehavior": "rename",
        "description": "description",
        "fileSize": file.size,
        "name": `${file.name}`
      };

      //Step1: Create the upload session      
      const requestURL = `/me/drive/root:/Uploads/${file.name}:/createUploadSession`;
      const uploadSessionRes = await props.graphClient.api(requestURL).post(payload);
      const uploadEndpoint: string = uploadSessionRes.uploadUrl;

      //Get file content      
      const fileBuffer = await file.arrayBuffer();

      // Maximum file chunk size
      const FILE_CHUNK_SIZE = 320 * 1024 // 0.32 MB;      

      //Total number of chunks for given file
      const NUM_CHUNKS = Math.floor(fileBuffer.byteLength / FILE_CHUNK_SIZE) + 1;

      //Counter for building progress bar
      let counter = 1;

      //Initial value of upload index
      let uploadIndex: number = 0;

      while (true) {

        //Get the current progress bar status
        const progressValue = parseFloat((counter / NUM_CHUNKS).toFixed(2));
        setPercentComplete(progressValue);

        //Calculate the end index
        let endIndex = uploadIndex + FILE_CHUNK_SIZE - 1;

        //Gets the slice
        let slice: ArrayBuffer;
        if (endIndex >= fileBuffer.byteLength) {
          endIndex = fileBuffer.byteLength - 1;
          slice = fileBuffer.slice(uploadIndex);
        } else {
          slice = fileBuffer.slice(uploadIndex, endIndex + 1);
        }

        //Upload file
        const headers = {
          'Content-Length': `${slice.byteLength}`,
          'Content-Range': `bytes ${uploadIndex}-${endIndex}/${fileBuffer.byteLength}`
        };

        LogHelper.info("", "", "Uploading chunk:" + `${uploadIndex}-${endIndex}`);

        const response = await props.graphClient.api(uploadEndpoint).headers(headers).put(slice);
        if (!response) {
          break;
        }
        if (response.nextExpectedRanges) {
          //Get the next expected range of the slice  
          uploadIndex = parseFloat(response.nextExpectedRanges[0]);
          counter++;
        } else {
          //if there is no next range then break the loop          
          //Gets the upoaded file response
          result = response
          break;
        }

      }

      setUploadedFile(result);
      setShowProgressBar(false);
      setIsUploadFinished(true);
    }

    catch (error) {
      console.log("Error in UploadLargeFileInChunks:", error);
      return null;
    }
  }



  return (
    <section className={`${styles.uploadLargeFile}`}>

      <div className="ms-Grid" dir="ltr">
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12">

            <Card type="inner" title="Resumable Upload of Large Files to Microsoft OneDrive" extra={<OneDriveIcon />}>
              <Dropzone
                label={"Drop Files here or click to browse"}
                style={{ minWidth: "505px" }}
                onChange={updateFiles}
                value={files}
                view={"list"}
                minHeight={"150px"}
                maxHeight={"30em"}
                clickable={!showProgressBar}
                header={!showProgressBar}
                disableScroll

              >
                {files.map((file) => (
                  <FileItem {...file}
                    onDelete={!showProgressBar ? handleDelete : null}
                    key={file.id}
                    info
                    alwaysActive
                    resultOnTooltip
                  />
                ))}
              </Dropzone>

              <div>
                {showProgressBar && <Progress percent={Number((percentComplete * 100).toFixed(0))} />}
              </div>

              <div className="ms-Grid-col ms-sm12" style={{ marginTop: '20px' }}>
                <DefaultButton
                  disabled={showProgressBar || files.length == 0}
                  className={styles.button}
                  text="Upload"
                  onClick={handleUploadFile}
                  iconProps={volume0Icon}
                />
              </div>
            </Card>

          </div>
          {isUploadFinished && showSucessAndReload(`${uploadedFile ? uploadedFile.name : "File"} has been successfully uploaded.`)}

        </div>
      </div>

    </section>
  );
}
