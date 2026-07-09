app.post("/api/excel/generate", (req, res) => {
    try {

        console.log("===== REQUEST BODY =====");
        console.log(JSON.stringify(req.body, null, 2));

        let data = req.body.data;
        const issueKey = req.body.issue_key || "TestCases";

        // String aaye to parse karo
        if (typeof data === "string") {
            data = JSON.parse(data);
        }

        // Single object aaye to array bana do
        if (!Array.isArray(data)) {
            data = [data];
        }

        console.log("Rows:", data.length);

        const ws = XLSX.utils.json_to_sheet(data);

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Test Cases");

        const buffer = XLSX.write(wb, {
            type: "buffer",
            bookType: "xlsx"
        });

        res.setHeader(
            "Content-Disposition",
            `attachment; filename=${issueKey}_TestCases.xlsx`
        );

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );

        return res.send(buffer);

    } catch (err) {

        console.error(err);

        return res.status(500).json({
            status: "error",
            message: err.message
        });

    }
});
