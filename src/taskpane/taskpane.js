/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady(info => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    }
});

export async function run() {
    /**
     * Insert your PowerPoint code here
     */
    const info = document.getElementById('info');
    info.innerHTML = "Getting selected slices...";
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, result => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            info.innerHTML = "Getting selected slides failed: " + result.error.message;
        } else {
            info.innerHTML = "Selected slide indices: " + result.value.slides.map(x => x.index).join();
        }
    });
}
