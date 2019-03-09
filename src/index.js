/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */



$(document).ready(() => {
    $('#run').click(run);
});
  
// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#success-msg').hide();
    $('#app-body').show();
};

async function run() {
    const setBodyHtml = async ({ value }) => {
        await Office.context.mailbox.item.body.setAsync(
            value,
            { coercionType: Office.CoercionType.Html }
        );
        $('#success-msg').show();
    }
    await Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, {}, setBodyHtml);
}