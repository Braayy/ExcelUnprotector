(() => {
    function downloadFile(filename, content) {
        const blobUrl = URL.createObjectURL(content);

        const el = document.createElement("a");
        el.href = blobUrl;
        el.download = filename;
        el.click();
    }

    async function removeSheetProtection(entry, zip) {
        const content = await entry.async("text");

        if (content.indexOf("sheetProtection") !== -1) {
            const unprotectedContent = content.replace(/<sheetProtection[^<>]*\/>/g, "");

            zip.file(entry.name, unprotectedContent);
        }
    }

    function unprotectWorkbook(file) {
        const reader = new FileReader();

        reader.onload = async (event) => {
            const zip = await JSZip.loadAsync(event.target.result);

            const promises = [];

            zip.forEach((relPath, entry) => {
                if (relPath.indexOf("xl/worksheets/") === -1) return;

                promises.push(removeSheetProtection(entry, zip));
            });

            await Promise.all(promises);

            const blob = await zip.generateAsync({ type: "blob" });

            downloadFile(`Desprotegido - ${file.name}`, blob);
        };

        reader.readAsArrayBuffer(file);
    }

    window.addEventListener("dragover", (event) => event.preventDefault());

    window.addEventListener("drop", (event) => {
        event.preventDefault();

        let files;

        if (event.dataTransfer.items) {
            files = Array.from(event.dataTransfer.items)
                .filter((item) => item.kind === "file")
                .map((item) => item.getAsFile());
        } else {
            files = event.dataTransfer.files;
        }

        unprotectWorkbook(files[0]);
    });

    document.querySelector("input").addEventListener("change", (event) => {
        unprotectWorkbook(event.currentTarget.files[0]);
    });
})();