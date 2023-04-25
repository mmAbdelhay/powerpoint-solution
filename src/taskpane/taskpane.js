/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("gotoslide").onclick = run;
    document.getElementById("startCountdown").onclick = startCountdown;
  }
});

export function run() {
  Office.context.document.goToByIdAsync(
    document.getElementById("slide-num").value,
    Office.GoToType.Index,
    function (asyncResult) {
      if (asyncResult.status == "failed") {
        showNotification("Error", asyncResult.error.message);
      }
    }
  );
}

// JavaScript
export function startCountdown() {
  // Get user input for the countdown time in minutes and seconds
  var minutesInput = parseInt(document.getElementById("minutes").value);
  var secondsInput = parseInt(document.getElementById("seconds").value);

  // Convert the countdown time to milliseconds
  var countdownTime = (minutesInput * 60 + secondsInput) * 1000;

  // Set the end time for the countdown
  var countdownEndTime = Date.now() + countdownTime;

  // Update the countdown every second
  var countdownInterval = setInterval(function () {
    // Get the current time
    var currentTime = Date.now();

    // Calculate the time remaining between now and the countdown end time
    var timeRemaining = countdownEndTime - currentTime;

    // Calculate the minutes and seconds remaining
    var minutes = Math.floor((timeRemaining % (1000 * 60 * 60)) / (1000 * 60));
    var seconds = Math.floor((timeRemaining % (1000 * 60)) / 1000);

    // Display the countdown in an HTML element
    document.getElementById("countdown").innerHTML = minutes + "m " + seconds + "s";

    // If the countdown is finished, clear the interval and display a message
    if (timeRemaining < 0) {
      clearInterval(countdownInterval);
      document.getElementById("countdown").innerHTML = "Time's up!";
      run();
    }
  }, 1000);
}
