function removeVietnameseTones(str) {
    if (!str) return "";
    return str.normalize('NFD')
              .replace(/[\u0300-\u036f]/g, '')
              .replace(/đ/g, 'd').replace(/Đ/g, 'D');
}

// ===== CLASS SINH VIÊN =====
class Student {
    constructor(fullName, studentId) {
        this.fullName = String(fullName).trim(); 
        this.studentId = String(studentId).trim(); 
    }

    getEmail() {
        // xoá phần có (LT)
        let name = this.fullName.replace(/\s*\(.*?\)\s*/g, "").trim();
        let nameParts = name.split(" ");
        let firstName = nameParts[nameParts.length - 1];

        let initials = "";
        for (let i = 0; i < nameParts.length - 1; i++) {
            if (nameParts[i][0]) initials += nameParts[i][0];
        }

        let cleanFirstName = removeVietnameseTones(firstName);
        let cleanInitials = removeVietnameseTones(initials);

        return (cleanFirstName + cleanInitials + "." + this.studentId).toLowerCase() + "@hvnh.edu.vn";
    }

    getCourse() {
        return this.studentId.substring(0, 2);
    }

    getFaculty() {
        let majorCode = this.studentId.substring(3, 6);

        switch (majorCode) {
            case "404": return "CNTT & KTS";
            case "408": return "KHDL";
            case "401": return "TC-NH";
            case "403": return "QTKD";
            case "405": return "KDQT";
            case "407": return "Kinh tế";
            case "406": return "Luật";
            case "751": return "Ngoại ngữ";
            case "402": return "Kế toán - Kiểm toán";
            default: return "Không xác định";
        }
    }

    getSumLast4() {
        let lastFour = this.studentId.slice(-4);
        let sum = 0;

        for (let i = 0; i < lastFour.length; i++) {
            if (!isNaN(lastFour[i])) sum += Number(lastFour[i]);
        }

        return sum;
    }
}
// ===== ĐỌC EXCEL =====
fetch("data.xlsx")
.then(res => res.arrayBuffer())
.then(data => {
    let workbook = XLSX.read(data, { type: "array" });
    let sheet = workbook.Sheets[workbook.SheetNames[0]];
    let rows = XLSX.utils.sheet_to_json(sheet);

    let html = "";

    for (let i = 0; i < rows.length; i++) {
        let keys = Object.keys(rows[i]);

        let rawName = rows[i][keys[2]];
        let rawId = rows[i][keys[1]];
        // bỏ qua dòng nếu thiếu tên hoặc mã sinh viên
        if (!rawName || !rawId) continue;

        // mỗi dòng là 1 sinh viên
        let sv = new Student(rawName, rawId);

        html += `
            <tr>
                <td>${sv.fullName}</td>
                <td>${sv.getEmail()}</td>
                <td>${sv.getCourse()}</td>
                <td>${sv.getFaculty()}</td>
                <td>${sv.getSumLast4()}</td>
            </tr>
        `;
    }

    document.getElementById("tableBody").innerHTML = html;
})
.catch(err => console.error("Lỗi fetch dữ liệu:", err));
