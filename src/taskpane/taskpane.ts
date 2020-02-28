/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import "zone.js"; // Required for Angular
import { platformBrowserDynamic } from "@angular/platform-browser-dynamic";
import AppModule from "./app/app.module";
/* global console, document, Office */

Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";

  localStorage.setItem("jsonFormatWithValues", JSON.stringify({
    questionCell: "B2",
    question: "QuestionGoesHere.",
    choiceOneCell: "B3",
    choiceOne: "Value One",
    choiceTwoCell: "B4",
    choiceTwo: "Value Two",
    choiceThreeCell: "B5",
    choiceThree: "Value Three",
    gradedCell: "B6",
    wrongAnswerColor: "red",
    rightAnswerColor: "green",
    rightAnswer: 5,
    formulae : "=SUM(2,3)"
  })
  );

  // Bootstrap the app
  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch(error => console.error(error));
};
