import React, { Component } from "react";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import PizZipUtils from "pizzip/utils/index.js";
import { saveAs } from "file-saver";
import DocxFile from "./resoureces/abc.docx";
import ImageModule from "docxtemplater-image-module-free";
import { testData } from "./testData";
import { testImage } from "./testData";

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

export const App = class App extends React.Component {
  render() {
    const generateDocument = async () => {
      try {
        const docxContent = await new Promise((resolve, reject) => {
          loadFile(DocxFile, function (error, content) {
            if (error) {
              reject(error);
            } else {
              resolve(content);
            }
          });
        });
    
        const imagePaths = testImage;
    
        const imagePromises = imagePaths.map((imagePath) => {
          return new Promise((resolve, reject) => {
            loadFile(imagePath, function (error, imageData) {
              if (error) {
                reject(error);
              } else {
                resolve(imageData);
              }
            });
          });
        });
    
        const imagesData = await Promise.all(imagePromises);
    
        generateDocWithImages(imagesData, docxContent);
      } catch (error) {
        console.error(error);
      }
    };

    const generateDocWithImages = (imagesData, docxContent) => {
      var opts = {};
      opts.centered = false;
      opts.getImage = function (tagValue, tagName) {
        const imgIndex = parseInt(tagValue) - 1; // Assuming image tags are numbered from 1 to n
        return imagesData[imgIndex]; // Return the resolved binary data for the respective image
      };
      opts.getSize = function (img, tagValue, tagName) {
        return [150, 150]; // Adjust the image size as needed
      };
    
      var imageModule = new ImageModule(opts);
    
      var zip = new PizZip(docxContent); // Use the direct binary content, not a Promise
      var doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        modules: [imageModule],
      });
    
      doc.setData({
        ...testData
      });
    
      try {
        // render the document (replace all occurrences of {first_name} by John, {last_name} by Doe, ...)
        doc.render();
      } catch (error) {
        // Handle the error
        console.error(error);
        throw error;
      }
    
      var out = doc.getZip().generate({
        type: "blob",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      }); // Output the document using Data-URI
      saveAs(out, "output.docx");
    };

    return (
      <div className="p-2">
        <button onClick={generateDocument}>Generate document</button>
      </div>
    );
  }
};
