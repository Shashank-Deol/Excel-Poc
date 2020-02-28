import { Component } from "@angular/core";
const template = require("./app.component.html");
/* global console, Excel, require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = "Welcome";
  answerByStudent: any;
  answerByAuthor = 5;
  jsonFormatWithValues: any;

  conditionToBeFulfilled() {
    if (this.answerByStudent === this.jsonFormatWithValues.rightAnswer) {
      return true;
    }
    else {
      return false;
    }
  }

  async run() {
    try {
      await Excel.run(async context => {

        this.jsonFormatWithValues = JSON.parse(localStorage.getItem("jsonFormatWithValues"));
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        // range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);

        /**
         * My code goes here for POC thing-
         */
        var mySheet = context.workbook.worksheets.getActiveWorksheet();

        if ((<HTMLInputElement>document.getElementById("test")).value === "openAssignment") {
          mySheet.protection.unprotect();
          var myRange = mySheet.getRange(this.jsonFormatWithValues.questionCell);
          myRange.values = [[this.jsonFormatWithValues.question]];
          myRange.format.autofitColumns();

          var myRangeForFirstAnswer = mySheet.getRange(this.jsonFormatWithValues.choiceOneCell);
          myRangeForFirstAnswer.values = [[this.jsonFormatWithValues.choiceOne]];
          myRange.format.autofitColumns();

          var myRangeForSecondAnswer = mySheet.getRange(this.jsonFormatWithValues.choiceTwoCell);
          myRangeForSecondAnswer.values = [[this.jsonFormatWithValues.choiceTwo]];
          myRange.format.autofitColumns();

          var myRangeForThirdAnswer = mySheet.getRange(this.jsonFormatWithValues.choiceThreeCell);
          myRangeForThirdAnswer.values = [[this.jsonFormatWithValues.choiceThree]];
          myRange.format.autofitColumns();

          var myRangeForGradedCell = mySheet.getRange(this.jsonFormatWithValues.gradedCell);
          myRangeForGradedCell.select();
          myRangeForGradedCell.format.fill.color = "yellow";
        }
        else if ((<HTMLInputElement>document.getElementById("test")).value === "takeAssignment") {
          // var mySheet = context.workbook.worksheets.getActiveWorksheet();
          // mySheet.protection.protect();
          var myRangeForTakeAssignment = mySheet.getRange(this.jsonFormatWithValues.gradedCell);

          myRangeForTakeAssignment.load("values");
          await context.sync();
          console.log(myRangeForTakeAssignment.values[0][0]);
          this.answerByStudent = myRangeForTakeAssignment.values[0][0];

          mySheet.protection.protect(null);
          myRangeForTakeAssignment.format.protection.load("locked");

          await context.sync();
          // localStorage.removeItem("jsonFormatWithValues");
          myRangeForTakeAssignment.format.protection.locked = false;

        }
        else {
          //  mySheet = context.workbook.worksheets.getActiveWorksheet();
          mySheet.protection.unprotect();
          var myRangeForPostReview = mySheet.getRange(this.jsonFormatWithValues.gradedCell);
          /**
           * Working
           */
          myRangeForPostReview.load("values");
          /**
           * TODO - 
           * locak a specific cell.
           */
          myRangeForPostReview.load('formulas');
          await context.sync();
          console.log(myRangeForPostReview.formulas);

          if (myRangeForPostReview.formulas[0][0] === this.jsonFormatWithValues.formulae || myRangeForPostReview.formulas[0][0] === this.jsonFormatWithValues.rightAnswer) {
            myRangeForPostReview.format.fill.color = this.jsonFormatWithValues.rightAnswerColor;
          }
          else {
            myRangeForPostReview.format.fill.color = this.jsonFormatWithValues.wrongAnswerColor;
          }

          // mySheet.protection.protect();
          myRangeForPostReview.format.protection.load("locked");
          await context.sync();
          myRangeForPostReview.format.protection.locked = true;
          await context.sync();
          mySheet.protection.protect();
          console.log(myRangeForPostReview.format.protection.locked);

        }

        await context.sync();

      });
    } catch (error) {
      console.error(error);
    }
  }
}
