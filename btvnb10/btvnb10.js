function removeVietnameseTones(str) {
    if (!str) return "";
    return str.normalize('NFD')
              .replace(/[\u0300-\u036f]/g, '')
              .replace(/đ/g, 'd').replace(/Đ/g, 'D');
}

fetch("data.xlsx")
.then(res => res.arrayBuffer())
.then(data => {
    let workbook = XLSX.read(data, { type: "array" });
    let sheet = workbook.Sheets[workbook.SheetNames[0]];
    let rows = XLSX.utils.sheet_to_json(sheet);
    console.log(rows);

    let html = "";

    for (let i = 0; i < rows.length; i++) {
        let keys = Object.keys(rows[i]);
        
        // Lấy dữ liệu thô từ Excel
        let rawName = rows[i][keys[2]]; // Cột Họ tên
        let rawId = rows[i][keys[1]];   // Cột Mã SV

        // KT: Nếu dòng trống thì bỏ qua để không bị lỗi
        if (!rawName || !rawId) continue;

        // Xóa nội dung trong ngoặc (LT) và khoảng trắng thừa
        let fullName = String(rawName).replace(/\s*\(.*?\)\s*/g, "").trim();
        let studentId = String(rawId).trim();

        // Tách tên và lấy chữ cái đầu
        let nameParts = fullName.split(" ");
        let firstName = nameParts[nameParts.length - 1];

        let initials = "";
        for (let j = 0; j < nameParts.length - 1; j++) {
            if (nameParts[j][0]) initials += nameParts[j][0];
        }

        // Lọc dấu họ tên
        let cleanFirstName = removeVietnameseTones(firstName);
        let cleanInitials = removeVietnameseTones(initials);
        let email = (cleanFirstName + cleanInitials + "." + studentId).toLowerCase() + "@hvnh.edu.vn";

        // Xử lý Khóa và Khoa
        let coursePeriod = studentId.substring(0, 2);
        let majorCode = studentId.substring(3, 6);
        let faculty = "";

        switch (majorCode) {
            case "404": faculty = "CNTT & KTS"; break;
            case "408": faculty = "KHDL"; break;
            case "401": faculty = "TC-NH"; break;
            case "403": faculty = "QTKD"; break;
            case "405": faculty = "KDQT"; break;
            case "407": faculty = "Kinh tế"; break;
            case "406": faculty = "Luật"; break;
            case "751": faculty = "Ngoại ngữ"; break;
            case "402": faculty = "Kế toán - Kiểm toán"; break;
            default: faculty = "Không xác định";
        }

        // Tính tổng 4 số cuối Mã SV
        let lastFour = studentId.slice(-4);
        let sum = 0;
        for (let k = 0; k < lastFour.length; k++) {
            if (!isNaN(lastFour[k])) sum += Number(lastFour[k]);
        }

        // IN RA BẢNG
        html += `
            <tr>
                <td>${fullName}</td>
                <td>${email}</td>
                <td>${coursePeriod}</td>
                <td>${faculty}</td>
                <td>${sum}</td>
            </tr>
        `;
    }
    document.getElementById("tableBody").innerHTML = html;
})
.catch(err => console.error("Lỗi fetch dữ liệu:", err));
