// --- UD Image Preview Logic ---
document.getElementById('ud-image').addEventListener('change', function (event) {
    const file = event.target.files[0];
    const preview = document.getElementById('ud-image-preview');
    delete preview.dataset.savedBase64;

    if (file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            preview.src = e.target.result;
            preview.style.display = 'block';
        }
        reader.readAsDataURL(file);
    } else {
        preview.src = '';
        preview.style.display = 'none';
    }
});

// --- Save UD Logic ---
const btnSaveUd = document.getElementById('btn-save-ud');
if (btnSaveUd) {
    btnSaveUd.addEventListener('click', async function () {
        const idInput = document.getElementById('ud-record-id');
        if (!idInput.value) idInput.value = generateId();

        const imageFile = document.getElementById('ud-image').files[0];
        const preview = document.getElementById('ud-image-preview');
        let imageBase64 = preview.dataset.savedBase64 || '';

        if (imageFile) {
            imageBase64 = await toBase64(imageFile);
            preview.dataset.savedBase64 = imageBase64;
        }

        const ps = document.getElementById('ud-ps').value.trim() || 'UD';
        const caseNo = document.getElementById('ud-case-no').value.trim() || 'Unknown';
        saveFileRecord({
            id: idInput.value,
            type: 'ud',
            title: `${ps} Case No-${caseNo}`,
            date: new Date().toLocaleString(),
            formData: serializeForm('ud-form'),
            imageBase64: imageBase64
        });
    });
}

// --- UD Form Submit Handler (Saves AND Generates PDF) ---
document.getElementById('ud-form').addEventListener('submit', async function (e) {
    e.preventDefault();

    const data = {
        ps: document.getElementById('ud-ps').value,
        pdType: document.getElementById('ud-pd-type').value,
        pdName: document.getElementById('ud-pd-name').value,
        caseNo: document.getElementById('ud-case-no').value,
        caseDate: formatDate(document.getElementById('ud-case-date').value),
        tracingDate: formatDate(document.getElementById('ud-tracing-date').value),
        tracingTime: document.getElementById('ud-tracing-time').value,
        tracingPlace: document.getElementById('ud-tracing-place').value,
        age: document.getElementById('ud-age-year').value ? document.getElementById('ud-age-year').value + ' Years' : '',
        sex: document.querySelector('input[name="ud-sex"]:checked')?.value || '',
        height: [
            document.getElementById('ud-height-feet').value ? document.getElementById('ud-height-feet').value + 'F.' : '',
            document.getElementById('ud-height-inch').value ? document.getElementById('ud-height-inch').value + ' inch' : ''
        ].filter(Boolean).join(' '),
        complexion: document.querySelector('input[name="ud-complexion"]:checked')?.value || '',
        build: document.querySelector('input[name="ud-build"]:checked')?.value || '',
        eyes: document.getElementById('ud-eyes').value,
        hair: document.querySelector('input[name="ud-hair"]:checked')?.value || '',
        apparel: document.getElementById('ud-apparel').value || '',
        specialMarks: document.getElementById('ud-special-marks').value || 'NIL',
        otherInfo: document.getElementById('ud-other-info').value || ''
    };

    const imageFile = document.getElementById('ud-image').files[0];
    const preview = document.getElementById('ud-image-preview');
    let imageBase64 = preview.dataset.savedBase64 || '';

    if (imageFile) {
        imageBase64 = await toBase64(imageFile);
        preview.dataset.savedBase64 = imageBase64;
    }

    // ---- Auto-save to localStorage on Generate ----
    const idInput = document.getElementById('ud-record-id');
    if (!idInput.value) idInput.value = generateId();

   await saveFileRecord({
        id: idInput.value,
        type: 'ud',
        title: `${data.ps || 'UD'} Case No-${data.caseNo || 'Unknown'}`,
        date: new Date().toLocaleString(),
        formData: serializeForm('ud-form'),
        imageBase64: imageBase64
    }, true);
    // ------------------------------------------------

    const page1 = generateUdTemplate1(data, imageBase64);
    const page2 = generateUdTemplate2(data, imageBase64);

    const pageBreak = `<div style="page-break-before: always; clear: both;"></div>`;
    const combinedHtml = page1 + pageBreak + page2;

    exportToPdf(combinedHtml, `${data.ps} UD Case ${(data.caseNo || '').replace(/\//g, '-')}`);
});

// UD Format 1: Director of Information
function generateUdTemplate1(data, imgBase64) {
    const imgTag = imgBase64 ? `<img src="${imgBase64}" width="120" height="150" style="background-color: #FFFFFF;" />` : '';

    return `
        <div style="font-family: 'Times New Roman', serif; font-size: 11pt; line-height: 1;">
            <div style="text-align: center; font-weight: bold;">
                Criminal Investigation Department<br>West Bengal<br>Bhabani Bhaban,<br>
                31, Belvedere Road, Alipore<br>Kolkata -700 027
            </div>
            <table width="100%" style="margin-top: 10px; margin-bottom: 10px;">
                <tr>
                    <td align="left">Memo No. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /MPB/CID/WB</td>
                    <td align="right">Date: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                </tr>
            </table>
            <p style="margin: 0;">To</p>
            <p style="margin: 0 0 15px 0;">
                The Director of Information<br>Department of Information & Cultural Affairs<br>
                Govt. of West Bengal<br>Nabanna, 9th Floor, 325, Sarat Chatterjee Road, Shibpur<br>Howrah 711 102
            </p>
            <p><strong>Ref: &nbsp;&nbsp;&nbsp; </strong> ${data.ps} (Dist - ${data.pdName} ${data.pdType}) U/D Case No. ${data.caseNo} date ${data.caseDate}.</p>
            <p style="text-align: justify;">In compliance with the guidelines issued by the Hon'ble Supreme Court of India on 14.11.2002 in Writ Petition (Crl.) No. 610 of 1996, in the matter of Horilal vs. Commissioner of Police, Delhi & Others, you are requested to circulate/publish the information of the following unidentified dead body in widely circulated newspapers-once each in English, Hindi, and Bengali.</p>
            <table width="100%">
                <tr><td width="35%">Name</td><td width="5%">:</td><td width="40%">Unknown ${data.sex} Deadbody</td><td rowspan="8" align="center" valign="top">${imgTag}</td></tr>
                <tr><td>Alias (es)</td><td>:</td><td></td></tr>
                <tr><td>Nickname</td><td>:</td><td></td></tr>
                <tr><td>Father's/Husband's Name</td><td>:</td><td></td></tr>
                <tr><td valign="top">Address</td><td valign="top">:</td><td></td></tr>
                <tr><td>Date of Tracing</td><td>:</td><td>${data.tracingDate}</td></tr>
                <tr><td>Time of Tracing</td><td>:</td>${data.tracingTime ? `<td>${data.tracingTime} hrs</td></tr>` : ''}
                <tr><td valign="top">Place of Tracing</td><td valign="top">:</td><td colspan="2">${data.tracingPlace}</td></tr>
            </table>
            <p style="text-decoration: underline; font-weight: bold; margin-top: 10px;">Descriptive Roll:</p>
            <table width="100%">
                <tr><td width="20%">Age:</td><td width="30%">${data.age}(Approx)</td><td width="20%">Sex:</td><td width="30%">${data.sex}</td></tr>
                <tr><td>Height:</td><td>${data.height} (Approx)</td><td>Complexion:</td><td>${data.complexion}</td></tr>
                <tr><td>Mother Tongue:</td><td></td><td>Build:</td><td>${data.build}</td></tr>
                <tr><td>Eyes:</td><td>${data.eyes}</td><td>Hair:</td><td>${data.hair}</td></tr>
                <tr><td>Mentally Challenge:</td><td></td><td valign="top">Wearing Apparels:</td><td valign="top">${data.apparel}</td></tr>
                <tr><td valign="top">Any Other Information:</td><td colspan="3">${data.otherInfo}</td></tr>
            </table>
            <p style="margin-top: 10px;">If identified, please contact:</p>
            <p style="margin-bottom: 30px;">
                Officer in Charge<br>Missing Persons Bureau<br>CID, West Bengal<br>
                Bhabani Bhaban, Kolkata - 700027<br>Phone No: 033-24506120
            </p>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr><td width="50%"></td><td width="50%" align="center" style="font-weight: bold;">Special Superintendent of Police (Spl)<br>Criminal Investigation Department<br>West Bengal</td></tr>
            </table>
        </div>
    `;
}

// UD Format 2: Hue & Cry Notice
function generateUdTemplate2(data, imgBase64) {
    const imgTag = imgBase64 ? `<img src="${imgBase64}" width="120" height="150" style="background-color: #FFFFFF;" />` : '';

    return `
        <div style="font-family: 'Times New Roman', serif; font-size: 11pt; line-height: 1.15;">
            <div style="text-align: center; font-weight: bold;">
                Criminal Investigation Department<br>West Bengal<br>Bhabani Bhaban,<br>
                31, Belvedere Road, Alipore<br>Kolkata -700 027
            </div>
             <p style="text-align: center; text-decoration: underline; font-weight: bold; font-size: 14pt; margin: 20px 0;">Hue & Cry Notice</p>
            <table width="100%" style="margin-top: 10px; margin-bottom: 10px;">
                <tr>
                    <td align="left">Org. No. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; /MPB/CID/WB</td>
                    <td align="right">Date: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                </tr>
            </table>
            <table width="100%" style="margin-bottom: 10px;">
                <tr>
                    <td width="8%" valign="top">To:</td>
                    <td width="92%" style="text-align: justify;">
                        DGSP of all States/UTs-w- DIGSP, CID of All States & UTs-w-CP, Kolkata Police -w- ALL CSP / SSP of West Bengal / All O/Cs of West Bengal -w- All OCs of Kolkata Police -w- All OCs of GRPSS / SCRB, West Bengal / PD, West Bengal.
                    </td>
                </tr>
                <tr><td valign="top">From:</td><td>Crime West Bengal.</td></tr>
                <tr><td valign="top">Ref:</td><td>${data.ps} (Dist - ${data.pdName} ${data.pdType}) U/D Case No. ${data.caseNo} date ${data.caseDate}.</td></tr>
                <tr><td valign="top">Sub:</td><td>Request for search for the identity of the recovered unidentified dead body.</td></tr>
            </table>
           
            <p style="text-align: justify;">The following unidentified dead body has been recovered from ${data.tracingPlace} under the jurisdiction of ${data.ps}, District - ${data.pdName}, West Bengal. The description of the recovered unidentified dead body is as follows:</p>
            <p style="text-decoration: underline; font-weight: bold;">Description of the recovered un-identified dead body:</p>
            <table width="100%">
                <tr><td width="35%">&#9679; Name</td><td width="5%">:</td><td width="40%">Unknown ${data.sex} Dead body</td><td rowspan="11" align="center" valign="top">${imgTag}</td></tr>
                <tr><td>&#9679; Father's/Husband's Name</td><td>:</td><td></td></tr>
                <tr><td valign="top">&#9679; Address</td><td valign="top">:</td><td></td></tr>
                <tr><td>&#9679; Age</td><td>:</td><td>${data.age}</td></tr>
                <tr><td>&#9679; Sex</td><td>:</td><td>${data.sex}</td></tr>
                <tr><td>&#9679; Date of Tracing</td><td>:</td><td>${data.tracingDate}</td></tr>
                <tr><td>&#9679; Height</td><td>:</td><td>${data.height} (Approx)</td></tr>
                <tr><td>&#9679; Complexion</td><td>:</td><td>${data.complexion}</td></tr>
                <tr><td>&#9679; Build</td><td>:</td><td>${data.build}</td></tr>
                <tr><td valign="top">&#9679; Special I'D Mark</td><td valign="top">:</td><td>${data.specialMarks}</td></tr>
                <tr><td valign="top">&#9679; Wearing Apparels</td><td valign="top">:</td><td>${data.apparel}</td></tr>
            </table>
            <p style="margin-top: 10px; text-align: justify;">You are requested to widely circulate the information regarding the recovered unidentified dead body within your jurisdiction.<br> Any information or clues about the recovered unidentified dead body, if available, may kindly be communicated to:</p>
            <p style="margin-bottom: 30px;">
                Officer in-Charge<br>Missing Persons Bureau<br>Bhabani Bhaban, Kolkata-700027<br>
                Phone No.: 033-24506120<br>Email: insp1mpbcid@policewb.gov.in
            </p>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr><td width="50%"></td><td width="50%" align="center" style="font-weight: bold;">Special Superintendent of Police (Spl)<br>Criminal Investigation Department<br>West Bengal</td></tr>
            </table>
        </div>
    `;
}