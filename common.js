// --- Tab Switching Logic ---
const btnMissing = document.getElementById('btn-missing');
const btnUd = document.getElementById('btn-ud');
const formMissing = document.getElementById('missing-form');
const formUd = document.getElementById('ud-form');

btnMissing.addEventListener('click', () => {
    btnMissing.classList.add('active');
    btnUd.classList.remove('active');
    formMissing.style.display = 'block';
    formUd.style.display = 'none';
});

btnUd.addEventListener('click', () => {
    btnUd.classList.add('active');
    btnMissing.classList.remove('active');
    formUd.style.display = 'block';
    formMissing.style.display = 'none';
});

// --- Shared Helper Functions ---

// Convert File to Base64 and Force JPEG (Strips Transparency)
function toBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        
        reader.onload = (event) => {
            const img = new Image();
            img.onload = () => {
                const canvas = document.createElement('canvas');
                canvas.width = img.width;
                canvas.height = img.height;
                const ctx = canvas.getContext('2d');
                
                ctx.fillStyle = '#FFFFFF';
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.drawImage(img, 0, 0);
                
                resolve(canvas.toDataURL('image/jpeg', 0.9)); 
            };
            img.onerror = error => reject(error);
            img.src = event.target.result;
        };
        reader.onerror = error => reject(error);
    });
}

// Convert date from yyyy-mm-dd to dd/mm/yyyy
function formatDate(dateString) {
    if (!dateString) return '';
    const parts = dateString.split('-');
    if (parts.length === 3) {
        return `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
    return dateString;
}

// Export HTML string to Word Document with Image Embedding (MHTML)
function exportToWord(htmlContent, filename) {
    let finalDocument = "";
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
                size: 595.3pt 841.9pt;
                margin: 0.5in 0.5in 0.5in 0.5in;
            }
            div.WordSection1 { page: WordSection1; }
            body { font-family: 'Times New Roman', serif; font-size: 11pt; line-height: 1.15; }
        </style>
    </head>
    <body>
        <div class="WordSection1">`;
    const postHtml = "</div></body></html>";

    if (match) {
        const mimeType = match[1];
        const base64Data = match[2];
        const globalBase64Regex = /data:image\/[^;]+;base64,[^"]+/g;
        const modifiedHtml = htmlContent.replace(globalBase64Regex, "embedded_image");
        const fullHtmlDocument = preHtml + modifiedHtml + postHtml;

        finalDocument = `MIME-Version: 1.0\nContent-Type: multipart/related; boundary="mht-boundary"\n\n--mht-boundary\nContent-Location: document.html\nContent-Type: text/html; charset="utf-8"\n\n${fullHtmlDocument}\n\n--mht-boundary\nContent-Location: embedded_image\nContent-Transfer-Encoding: base64\nContent-Type: ${mimeType}\n\n${base64Data}\n--mht-boundary--`;
    } else {
        finalDocument = preHtml + htmlContent + postHtml;
    }

    const blob = new Blob(['\ufeff', finalDocument], { type: 'application/msword' });
    const downloadLink = document.createElement("a");
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = filename.endsWith('.doc') ? filename : filename + '.doc';
    
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
}