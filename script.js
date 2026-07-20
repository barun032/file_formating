document.getElementById('missing-form').addEventListener('submit', async function (e) {
    e.preventDefault();

    // Gather basic data
    const data = {
        ps: document.getElementById('ps').value,
        pdType: document.getElementById('pd-type').value,
        pdName: document.getElementById('pd-name').value,
        gdeNo: document.getElementById('gde-no').value,
        gdeDate: formatDate(document.getElementById('gde-date').value),
        caseNo: document.getElementById('case-no').value,
        caseDate: formatDate(document.getElementById('case-date').value),
        caseSection: document.getElementById('case-section').value,
        name: document.getElementById('missing-name').value,
        alias: document.getElementById('alias').value,
        nickname: document.getElementById('nickname').value,
        guardianType: document.getElementById('guardian-type').value,
        guardianName: document.getElementById('guardian-name').value,
        address: document.getElementById('address').value,
        missingDate: formatDate(document.getElementById('missing-date').value),
        missingTime: document.getElementById('missing-time').value,
        age: [
            document.getElementById('age-year').value ? document.getElementById('age-year').value + ' Years' : '',
            document.getElementById('age-month').value ? document.getElementById('age-month').value + ' Months' : ''
        ].filter(Boolean).join(' '),
        sex: document.querySelector('input[name="sex"]:checked')?.value || '',
        height: [
            document.getElementById('height-feet').value ? document.getElementById('height-feet').value + ' Fit.' : '',
            document.getElementById('height-inch').value ? document.getElementById('height-inch').value + ' Inch' : ''
        ].filter(Boolean).join(' '),
        complexion: document.querySelector('input[name="complexion"]:checked')?.value || '',
        language: document.getElementById('language').value,
        build: document.querySelector('input[name="build"]:checked')?.value || '',
        eyes: document.getElementById('eyes').value,
        hair: document.querySelector('input[name="hair"]:checked')?.value || '',
        mental: document.querySelector('input[name="mental"]:checked')?.value || '',
        apparel: document.getElementById('apparel').value || '',
        specialMarks: document.getElementById('special-marks') ? document.getElementById('special-marks').value : ''
    };

    // Handle Image upload
    const imageFile = document.getElementById('image').files[0];
    let imageBase64 = '';

    if (imageFile) {
        imageBase64 = await toBase64(imageFile);
    }

    // Generate individual templates
    const page1 = generateTemplate1(data, imageBase64);
    const page2 = generateTemplate2(data, imageBase64);
    const page3 = generateTemplate3(data, imageBase64);
    const page4 = generateTemplate4(data, imageBase64);

    // Combine them with Word-compatible page breaks
    const pageBreak = `<br clear="all" style="page-break-before:always" />`;
    const combinedHtml = page1 + pageBreak + page2 + pageBreak + page3 + pageBreak + page4;

    // Export as a single document
    exportToWord(combinedHtml, `${data.name}_${data.age}`);
});

// Helper: Convert File to Base64
function toBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result);
        reader.onerror = error => reject(error);
    });
}

// --- Image Preview Logic ---
document.getElementById('image').addEventListener('change', function (event) {
    const file = event.target.files[0];
    const preview = document.getElementById('image-preview');

    if (file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            preview.src = e.target.result;
            preview.style.display = 'block'; // Show the image
        }
        reader.readAsDataURL(file);
    } else {
        preview.src = '';
        preview.style.display = 'none'; // Hide if no file is selected
    }
});

// Helper: Convert date from yyyy-mm-dd to dd/mm/yyyy
function formatDate(dateString) {
    if (!dateString) return '';
    const parts = dateString.split('-');
    if (parts.length === 3) {
        return `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
    return dateString;
}

// Helper: Export HTML string to Word Document with Image Embedding (MHTML)
function exportToWord(htmlContent, filename) {
    let finalDocument = "";

    // Check if the HTML contains a base64 image (we use a global regex now to catch multiple instances if needed)
    const base64Regex = /data:(image\/[^;]+);base64,([^"]+)/;
    const match = htmlContent.match(base64Regex);

    const preHtml = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
        <meta charset='utf-8'>
        <title>Export</title>
        <!--[if gte mso 9]>
        <xml>
            <w:WordDocument>
                <w:View>Print</w:View>
                <w:Zoom>100</w:Zoom>
                <w:DoNotOptimizeForBrowser/>
            </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
            @page WordSection1 {
                size: 595.3pt 841.9pt; /* A4 Paper Size */
                margin: 0.5in 0.5in 0.5in 0.5in; /* Narrow Margins (0.5 inch) */
            }
            div.WordSection1 { 
                page: WordSection1; 
            }
            body {
                font-family: 'Times New Roman', serif;
                font-size: 11pt;
                line-height: 1.15;
            }
        </style>
    </head>
    <body>
        <div class="WordSection1">`;
    const postHtml = "</div></body></html>";

    if (match) {
        // We have an image! Extract the type and the raw data.
        const mimeType = match[1];
        const base64Data = match[2];

        // Replace ALL instances of the base64 string in the HTML with a local reference Word can understand
        const globalBase64Regex = /data:image\/[^;]+;base64,[^"]+/g;
        const modifiedHtml = htmlContent.replace(globalBase64Regex, "embedded_image");
        const fullHtmlDocument = preHtml + modifiedHtml + postHtml;

        // Create an MHTML package so MS Word loads the image natively
        finalDocument = `MIME-Version: 1.0
Content-Type: multipart/related; boundary="mht-boundary"

--mht-boundary
Content-Location: document.html
Content-Type: text/html; charset="utf-8"

${fullHtmlDocument}

--mht-boundary
Content-Location: embedded_image
Content-Transfer-Encoding: base64
Content-Type: ${mimeType}

${base64Data}
--mht-boundary--`;
    } else {
        // No image provided by the user, export as standard HTML
        finalDocument = preHtml + htmlContent + postHtml;
    }

    const blob = new Blob(['\ufeff', finalDocument], { type: 'application/msword' });
    const downloadLink = document.createElement("a");
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = filename;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
}

// ----------------------------------------------------
// OFFICIAL CID WEST BENGAL TEMPLATES
// ----------------------------------------------------

// Format 1: Director of Information (Newspapers)
function generateTemplate1(data, imgBase64) {
    const imgTag = imgBase64 ? `<img src="${imgBase64}" width="120" height="150" />` : '';

    // Dynamic logic for Case or GDE NO
    const caseOrGde = data.caseNo ? `Case NO.-${data.caseNo} Date-${data.caseDate} ${data.caseSection}` : `GDE NO.-${data.gdeNo} Date-${data.gdeDate}`;

    return `
        <div style="font-family: 'Times New Roman', serif; font-size: 11pt; line-height: 1.15;">
            <div style="text-align: center; font-weight: bold;">
                Criminal Investigation Department<br>
                West Bengal<br>
                Bhabani Bhaban,<br>
                31, Belvedere Road, Alipore<br>
                Kolkata-700 027
            </div>
            
            <table width="100%" style="margin-top: 15px; margin-bottom: 15px;">
                <tr>
                    <td align="left">Memo No. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /MPB/CID/WB</td>
                    <td align="right">Date: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                </tr>
            </table>
            
            <p style="margin: 0;">To</p>
            <p style="margin: 0 0 15px 0;">
                The Director of Information<br>
                Department of Information & Cultural Affairs<br>
                Govt. of West Bengal<br>
                Nabanna, 9th Floor, 325, Sarat Chatterjee Road, Shibpur<br>
                Howrah 711 102
            </p>
            
            <p><strong>Ref: &nbsp;&nbsp;&nbsp; </strong> ${data.ps} PS (${data.pdName} ${data.pdType}) ${caseOrGde}.</p>
            
            <p style="text-align: justify;">In compliance with the guidelines issued by the Hon'ble Supreme Court of India on 14.11.2002 in Writ Petition (Crl.) No. 610 of 1996, in the matter of Horilal vs. Commissioner of Police, Delhi & Others, you are requested to circulate/publish the information of the following missing person in widely circulated newspapers-once-each in English, Hindi, and Bengali.</p>
            
            <table width="100%">
                <tr>
                    <td width="35%">Name</td><td width="5%">:</td><td width="40%">${data.name}</td>
                    <td rowspan="7" align="center" valign="top">${imgTag}</td>
                </tr>
                <tr><td>Alias (es)</td><td>:</td><td>${data.alias}</td></tr>
                <tr><td>Nickname</td><td>:</td><td>${data.nickname}</td></tr>
                <tr><td>${data.guardianType}</td><td>:</td><td>${data.guardianName}</td></tr>
                <tr><td valign="top">Address</td><td valign="top">:</td><td>${data.address}</td></tr>
                <tr><td>Date of Missing</td><td>:</td><td>${data.missingDate}</td></tr>
                <tr><td>Time of Missing</td><td>:</td><td>${data.missingTime} hrs</td></tr>
                <tr><td valign="top">Place of Missing</td><td valign="top">:</td><td colspan="2">Residential area under ${data.ps} PS.</td></tr>
            </table>
            
            <p style="text-decoration: underline; font-weight: bold; margin-top: 15px;">Descriptive Roll:</p>
            <table width="100%">
                <tr>
                    <td width="20%">Age:</td><td width="30%">${data.age}(Approx)</td>
                    <td width="20%">Sex:</td><td width="30%">${data.sex}</td>
                </tr>
                <tr>
                    <td>Height:</td><td>${data.height}(Approx)</td>
                    <td>Complexion:</td><td>${data.complexion}</td>
                </tr>
                <tr>
                    <td>Mother Tongue:</td><td>${data.language}</td>
                    <td>Build:</td><td>${data.build}</td>
                </tr>
                <tr>
                    <td>Eyes:</td><td>${data.eyes}</td>
                    <td>Hair:</td><td>${data.hair}</td>
                </tr>
                <tr>
                    <td>Mentally Challenge:</td><td>${data.mental}</td>
                    <td>Wearing Apparels:</td><td>${data.apparel}</td>
                </tr>
                 <tr>
                  <td>Special I'd Mark:</td><td>${data.specialMarks}</td>
                </tr>
            </table>
            
            <p style="margin-top: 20px;">If traced, please contact:</p>
            <p style="margin-bottom: 30px;">
                Officer in Charge<br>
                Missing Persons Bureau<br>
                CID, West Bengal<br>
                Bhabani Bhaban, Kolkata - 700027<br>
                Phone No: 033-24506120
            </p>
            
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td width="50%"></td>
                    <td width="50%" align="center" style="font-weight: bold;">
                        Special Superintendent of Police (Spl)<br>
                        Criminal Investigation Department<br>
                        West Bengal
                    </td>
                </tr>
            </table>
        </div>
    `;
}

// Format 2: Doordarshan Kendra (Telecast)
function generateTemplate2(data, imgBase64) {
    const imgTag = imgBase64 ? `<img src="${imgBase64}" width="120" height="150" />` : '';

    // Dynamic logic for Case or GDE NO
    const caseOrGde = data.caseNo ? `Case No.-${data.caseNo} Date-${data.caseDate} ${data.caseSection}` : `GDE No.-${data.gdeNo} Date-${data.gdeDate}`;

    return `
        <div style="font-family: 'Times New Roman', serif; font-size: 11pt; line-height: 1.15;">
            <div style="text-align: center; font-weight: bold;">
                Criminal Investigation Department<br>
                West Bengal<br>
                Bhabani Bhaban,<br>
                31, Belvedere Road, Alipore<br>
                Kolkata-700 027
            </div>
            
            <table width="100%" style="margin-top: 15px; margin-bottom: 15px;">
                <tr>
                    <td align="left">Memo No. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /MPB/CID/WB</td>
                    <td align="right">Date: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                </tr>
            </table>
            
            <p style="margin: 0;">To</p>
            <p style="margin: 0 0 15px 0;">
                The Deputy Director General (Engg.)<br>
                Doordarshan Kendra, Kolkata<br>
                Doordarshan Bhavan,<br>
                18/3 Uday Shankar Sarani,<br>
                Golf Green,<br>
                Kolkata-700 095
            </p>
            
            <p><strong>Ref: &nbsp;&nbsp;&nbsp; </strong> ${data.ps} PS (${data.pdName} ${data.pdType}) ${caseOrGde}.</p>
            
            <p style="text-align: justify;">In compliance with the guidelines issued by the Hon'ble Supreme Court of India on 14.11.2002 in Writ Petition (Crl.) No. 610 of 1996, in the matter of Horilal vs. Commissioner of Police, Delhi & Others, you are requested to telecast the information of the following missing person through Doordarshan.</p>
            
            <table width="100%">
                <tr>
                    <td width="35%">Name</td><td width="5%">:</td><td width="40%">${data.name}</td>
                    <td rowspan="7" align="center" valign="top">${imgTag}</td>
                </tr>
                <tr><td>Alias (es)</td><td>:</td><td>${data.alias}</td></tr>
                <tr><td>Nickname</td><td>:</td><td>${data.nickname}</td></tr>
                <tr><td>${data.guardianType}</td><td>:</td><td>${data.guardianName}</td></tr>
                <tr><td valign="top">Address</td><td valign="top">:</td><td>${data.address}</td></tr>
                <tr><td>Date of Missing</td><td>:</td><td>${data.missingDate}</td></tr>
                <tr><td>Time of Missing</td><td>:</td><td>${data.missingTime} hrs</td></tr>
                <tr><td valign="top">Place of Missing</td><td valign="top">:</td><td colspan="2">Residential area under ${data.ps} PS.</td></tr>
            </table>
            
            <p style="text-decoration: underline; font-weight: bold; margin-top: 15px;">Descriptive Roll:</p>
            <table width="100%">
                <tr>
                    <td width="20%">Age:</td><td width="30%">${data.age}(Approx)</td>
                    <td width="20%">Sex:</td><td width="30%">${data.sex}</td>
                </tr>
                <tr>
                    <td>Height:</td><td>${data.height}(Approx)</td>
                    <td>Complexion:</td><td>${data.complexion}</td>
                </tr>
                <tr>
                    <td>Mother Tongue:</td><td>${data.language}</td>
                    <td>Build:</td><td>${data.build}</td>
                </tr>
                <tr>
                    <td>Eyes:</td><td>${data.eyes}</td>
                    <td>Hair:</td><td>${data.hair}</td>
                </tr>
                <tr>
                    <td>Mentally Challenge:</td><td>${data.mental}</td>
                    <td>Wearing Apparels:</td><td>${data.apparel}</td>
                </tr>
                 <tr>
                  <td>Special I'd Mark:</td><td>${data.specialMarks}</td>
                </tr>
            </table>
            
            <p style="margin-top: 20px;">Any information or clues about the missing person, if available, may kindly be communicated to:</p>
            <p style="margin-bottom: 30px;">
                Officer in Charge<br>
                Missing Persons Bureau<br>
                CID, West Bengal<br>
                Bhabani Bhaban, Kolkata - 700027<br>
                Phone No: 033-24506120<br>
                Email: insp1mpbcid@policewb.gov.in
            </p>
            
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td width="50%"></td>
                    <td width="50%" align="center" style="font-weight: bold;">
                        Special Superintendent of Police (Spl)<br>
                        Criminal Investigation Department<br>
                        West Bengal
                    </td>
                </tr>
            </table>
        </div>
    `;
}

// Format 3: Akashvani Kolkata (Broadcast)
function generateTemplate3(data, imgBase64) {
    const imgTag = imgBase64 ? `<img src="${imgBase64}" width="120" height="150" />` : '';

    // Dynamic logic for Case or GDE NO
    const caseOrGde = data.caseNo ? `Case No.-${data.caseNo} Date-${data.caseDate} ${data.caseSection}` : `GDE No.-${data.gdeNo} Date-${data.gdeDate}`;

    return `
        <div style="font-family: 'Times New Roman', serif; font-size: 11pt; line-height: 1.15;">
            <div style="text-align: center; font-weight: bold;">
                Criminal Investigation Department<br>
                West Bengal<br>
                Bhabani Bhaban,<br>
                31, Belvedere Road, Alipore<br>
                Kolkata-700 027
            </div>
            
            <table width="100%" style="margin-top: 15px; margin-bottom: 15px;">
                <tr>
                    <td align="left">Memo No. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /MPB/CID/WB</td>
                    <td align="right">Date: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                </tr>
            </table>
            
            <p style="margin: 0;">To</p>
            <p style="margin: 0 0 15px 0;">
                The Deputy Director General (Engg.)<br>
                Akashvani Kolkata<br>
                Akashvani Bhavan<br>
                Gostho Paul Sarani, Near Eden Gardens,<br>
                Kolkata 700 001
            </p>
            
            <p><strong>Ref: &nbsp;&nbsp;&nbsp; </strong> ${data.ps} PS (${data.pdName} ${data.pdType}) ${caseOrGde}.</p>
            
            <p style="text-align: justify;">In compliance with the guidelines issued by the Hon'ble Supreme Court of India on 14.11.2002 in Writ Petition (Crl.) No. 610 of 1996, in the matter of Horilal vs. Commissioner of Police, Delhi & Others, you are requested to broadcast the information of the following missing person through Akashvani.</p>
            
            <table width="100%">
                <tr>
                    <td width="35%">Name</td><td width="5%">:</td><td width="40%">${data.name}</td>
                    <td rowspan="7" align="center" valign="top">${imgTag}</td>
                </tr>
                <tr><td>Alias (es)</td><td>:</td><td>${data.alias}</td></tr>
                <tr><td>Nickname</td><td>:</td><td>${data.nickname}</td></tr>
                <tr><td>${data.guardianType}</td><td>:</td><td>${data.guardianName}</td></tr>
                <tr><td valign="top">Address</td><td valign="top">:</td><td>${data.address}</td></tr>
                <tr><td>Date of Missing</td><td>:</td><td>${data.missingDate}</td></tr>
                <tr><td>Time of Missing</td><td>:</td><td>${data.missingTime} hrs</td></tr>
                <tr><td valign="top">Place of Missing</td><td valign="top">:</td><td colspan="2">Residential area under ${data.ps} PS.</td></tr>
            </table>
            
            <p style="text-decoration: underline; font-weight: bold; margin-top: 15px;">Descriptive Roll:</p>
            <table width="100%">
                <tr>
                    <td width="20%">Age:</td><td width="30%">${data.age}(Approx)</td>
                    <td width="20%">Sex:</td><td width="30%">${data.sex}</td>
                </tr>
                <tr>
                    <td>Height:</td><td>${data.height}(Approx)</td>
                    <td>Complexion:</td><td>${data.complexion}</td>
                </tr>
                <tr>
                    <td>Mother Tongue:</td><td>${data.language}</td>
                    <td>Build:</td><td>${data.build}</td>
                </tr>
                <tr>
                    <td>Eyes:</td><td>${data.eyes}</td>
                    <td>Hair:</td><td>${data.hair}</td>
                </tr>
                <tr>
                    <td>Mentally Challenge:</td><td>${data.mental}</td>
                    <td>Wearing Apparels:</td><td>${data.apparel}</td>
                </tr>
                 <tr>
                    <td>Special I'd Mark:</td><td>${data.specialMarks}</td>
                </tr>
            </table>
            
            <p style="margin-top: 20px;">Any information or clues about the missing person, if available, may kindly be communicated to:</p>
            <p style="margin-bottom: 30px;">
                Officer in-Charge<br>
                Missing Persons Bureau<br>
                CID, West Bengal<br>
                Bhabani Bhaban, Kolkata-700027<br>
                Phone No.: 033-24506120<br>
                Email: insp1mpbcid@policewb.gov.in
            </p>
            
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td width="50%"></td>
                    <td width="50%" align="center" style="font-weight: bold;">
                        Special Superintendent of Police (Spl)<br>
                        Criminal Investigation Department<br>
                        West Bengal
                    </td>
                </tr>
            </table>
        </div>
    `;
}

// Format 4: Hue & Cry Notice
function generateTemplate4(data, imgBase64) {
    const imgTag = imgBase64 ? `<img src="${imgBase64}" width="120" height="150" />` : '';

    // Dynamic logic for Case or GDE NO
    const caseOrGde = data.caseNo ? `Case No.-${data.caseNo} Date-${data.caseDate} ${data.caseSection}` : `GDE No.-${data.gdeNo} Date-${data.gdeDate}`;

    return `
        <div style="font-family: 'Times New Roman', serif; font-size: 11pt; line-height: 1.15;">
            <div style="text-align: center; font-weight: bold;">
                Criminal Investigation Department<br>
                West Bengal<br>
                Bhabani Bhaban,<br>
                31, Belvedere Road, Alipore<br>
                Kolkata -700 027<br><br>
                <span style="text-decoration: underline; font-size: 14pt;">Hue & Cry Notice</span>
            </div>
            
            <table width="100%" style="margin-top: 15px; margin-bottom: 15px;">
                <tr>
                    <td align="left">Org. No. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /MPB/CID/WB</td>
                    <td align="right">Date: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                </tr>
            </table>
            
            <table width="100%" style="margin-bottom: 15px;">
                <tr>
                    <td width="8%" valign="top">To:</td>
                    <td width="92%" style="text-align: justify;">
                        DGSP of all States / UTs-w- DIGsP, CID of All States & UTs-w- CP, Kolkata Police -w-<br>
                        ALL CSP/SSP of West Bengal / All O/Cs of West Bengal-w- All OCs of Kolkata Police -w-<br>
                        All OCs of GRPSS/SCRB, West Bengal / PD, West Bengal.
                    </td>
                </tr>
                <tr>
                    <td valign="top">From:</td>
                    <td>Crime West Bengal.</td>
                </tr>
                <tr>
                    <td valign="top">Ref:</td>
                    <td>${data.ps} PS (${data.pdName} ${data.pdType}) ${caseOrGde}.</td>
                </tr>
                <tr>
                    <td valign="top">Sub:</td>
                    <td>Request for search for missing person.</td>
                </tr>
            </table>
            
            <p style="text-align: justify;">The following person has been reported missing from residential area under the jurisdiction of ${data.ps} PS, District - ${data.pdName}, West Bengal. The description of the missing person is as follows:</p>
            
            <p style="text-decoration: underline; font-weight: bold;">Description of the Missing Person:</p>
            
            <table width="100%">
                <tr>
                    <td width="35%">&#9679; Name</td><td width="5%">:</td><td width="40%">${data.name}</td>
                    <td rowspan="9" align="center" valign="top">${imgTag}</td>
                </tr>
                <tr><td>&#9679; ${data.guardianType}</td><td>:</td><td>${data.guardianName}</td></tr>
                <tr><td valign="top">&#9679; Address</td><td valign="top">:</td><td>${data.address}</td></tr>
                <tr><td>&#9679; Age</td><td>:</td><td>${data.age}</td></tr>
                <tr><td>&#9679; Sex</td><td>:</td><td>${data.sex}</td></tr>
                <tr><td>&#9679; Height</td><td>:</td><td>${data.height}</td></tr>
                <tr><td>&#9679; Complexion</td><td>:</td><td>${data.complexion}</td></tr>
                <tr><td>&#9679; Build</td><td>:</td><td>${data.build}</td></tr>
                <tr><td valign="top">&#9679; Wearing Apparels</td><td valign="top">:</td><td>${data.apparel}</td></tr>
                <tr><td valign="top">&#9679; Special I'd Mark</td><td valign="top">:</td><td>${data. specialMarks}</td></tr>
            </table>
            
            <p style="margin-top: 15px; text-align: justify;">You are requested to widely circulate the information regarding the missing person within your jurisdiction.<br> Any information or clues about the missing person, if available, may kindly be communicated to:</p>

            <p style="margin-bottom: 30px;">
                Officer in-Charge<br>
                Missing Persons Bureau<br>
                CID, West Bengal<br>
                Bhabani Bhaban, Kolkata-700027<br>
                Phone No.: 033-24506120<br>
                Email: insp1mpbcid@policewb.gov.in
            </p>
            
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td width="50%"></td>
                    <td width="50%" align="center" style="font-weight: bold;">
                        Special Superintendent of Police (Spl)<br>
                        Criminal Investigation Department<br>
                        West Bengal
                    </td>
                </tr>
            </table>
        </div>
    `;
}