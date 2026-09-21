import pandas as pd
from pathlib import Path

def generate_default_template(output_path: str = "templates/Config_Template.xlsx"):
    """Tự động sinh file Excel mẫu chuẩn cho hệ thống MCDM (gọn gàng, linh hoạt)."""
    
    # 1. Sheet Factors & Alternatives
    df_factors = pd.DataFrame([
        {"Type": "Factor", "Code": "C1", "Name": "Chi phí đầu tư (Cost)", "Description": "Tổng chi phí triển khai dự án"},
        {"Type": "Factor", "Code": "C2", "Name": "Thời gian hoàn thành (Time)", "Description": "Tiến độ thực hiện"},
        {"Type": "Factor", "Code": "C3", "Name": "Chất lượng kỹ thuật (Quality)", "Description": "Độ ổn định, đáp ứng kỹ thuật"},
        {"Type": "Factor", "Code": "C4", "Name": "Độ tin cậy của đối tác (Reliability)", "Description": "Uy tín và năng lực nhà cung cấp"},
        {"Type": "Alternative", "Code": "A1", "Name": "Phương án 1 (Vendor X)", "Description": "Thuê trọn gói giải pháp ngoài"},
        {"Type": "Alternative", "Code": "A2", "Name": "Phương án 2 (In-house)", "Description": "Tự phát triển nội bộ"},
        {"Type": "Alternative", "Code": "A3", "Name": "Phương án 3 (Hybrid)", "Description": "Thuê chuyên gia cố vấn và làm nội bộ"}
    ])

    # 2. Sheet Thang đo thông thường (Không có Z): Đúng 3 cột, Value điền tự do
    df_scale_standard = pd.DataFrame([
        {"Linguistic_Term": "Very Low", "Code": "VL", "Value": "0, 0, 0.25"},
        {"Linguistic_Term": "Low", "Code": "L", "Value": "0, 0.25, 0.5"},
        {"Linguistic_Term": "Medium", "Code": "M", "Value": "0.25, 0.5, 0.75"},
        {"Linguistic_Term": "High", "Code": "H", "Value": "0.5, 0.75, 1.0"},
        {"Linguistic_Term": "Very High", "Code": "VH", "Value": "0.75, 1.0, 1.0"}
    ])

    # 3. Sheet Thang đo Z-number: 6 cột (ghép 2 bộ 3 cột: Đánh giá A & Chắc chắn B)
    df_scale_z_number = pd.DataFrame([
        {
            "A_Linguistic_Term": "Very Low", "A_Code": "VL", "A_Value": "0, 0, 0.25",
            "B_Linguistic_Term": "Very Unlikely", "B_Code": "VU", "B_Value": "0, 0, 0.2"
        },
        {
            "A_Linguistic_Term": "Low", "A_Code": "L", "A_Value": "0, 0.25, 0.5",
            "B_Linguistic_Term": "Unlikely", "B_Code": "U", "B_Value": "0.1, 0.3, 0.5"
        },
        {
            "A_Linguistic_Term": "Medium", "A_Code": "M", "A_Value": "0.25, 0.5, 0.75",
            "B_Linguistic_Term": "Neutral", "B_Code": "N", "B_Value": "0.3, 0.5, 0.7"
        },
        {
            "A_Linguistic_Term": "High", "A_Code": "H", "A_Value": "0.5, 0.75, 1.0",
            "B_Linguistic_Term": "Likely", "B_Code": "Lk", "B_Value": "0.5, 0.7, 0.9"
        },
        {
            "A_Linguistic_Term": "Very High", "A_Code": "VH", "A_Value": "0.75, 1.0, 1.0",
            "B_Linguistic_Term": "Very Likely", "B_Code": "VLk", "B_Value": "0.8, 1.0, 1.0"
        }
    ])

    # 4. Sheet Chuyên gia: 3 cột
    df_experts = pd.DataFrame([
        {"Expert_Name": "GS. Nguyen Van A", "Username": "expert01", "Password": "123"},
        {"Expert_Name": "TS. Tran Thi B", "Username": "expert02", "Password": "123"},
        {"Expert_Name": "ThS. Le Van C", "Username": "expert03", "Password": "123"}
    ])

    # Ghi ra file Excel
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_factors.to_excel(writer, sheet_name="Factors_Alternatives", index=False)
        df_scale_standard.to_excel(writer, sheet_name="Scale_Standard", index=False)
        df_scale_z_number.to_excel(writer, sheet_name="Scale_Z_Number", index=False)
        df_experts.to_excel(writer, sheet_name="Experts", index=False)
        
    print(f"[OK] Đã tạo template mẫu thành công tại: {output_path}")

from core.models import ProjectConfig, Factor, ScaleItem, Expert

def parse_config_excel(file_path: str, config: ProjectConfig) -> ProjectConfig:
    """Đọc file Excel và nạp dữ liệu vào đối tượng ProjectConfig."""
    xls = pd.ExcelFile(file_path)
    config.excel_path = file_path
    
    # 1. Đọc sheet Factors_Alternatives
    if "Factors_Alternatives" in xls.sheet_names:
        df_fa = pd.read_excel(xls, "Factors_Alternatives").fillna("")
        config.factors = []
        config.alternatives = []
        for _, row in df_fa.iterrows():
            item_type = str(row.get("Type", "Factor")).strip()
            item = Factor(
                type=item_type,
                code=str(row.get("Code", "")).strip(),
                name=str(row.get("Name", "")).strip(),
                description=str(row.get("Description", "")).strip()
            )
            if item_type.lower() == "alternative":
                config.alternatives.append(item)
            else:
                config.factors.append(item)

    # 2. Đọc Sheet thang đo
    if config.scale_type == "Category":
        if not config.use_z_number:
            # Thang đo thông thường (3 cột)
            sheet_name = "Scale_Standard" if "Scale_Standard" in xls.sheet_names else xls.sheet_names[1]
            df_scale = pd.read_excel(xls, sheet_name).fillna("")
            config.standard_scales = []
            for _, row in df_scale.iterrows():
                config.standard_scales.append(ScaleItem(
                    linguistic=str(row.iloc[0]).strip(),
                    code=str(row.iloc[1]).strip(),
                    value=row.iloc[2]
                ))
        else:
            # Thang đo Z-number (6 cột)
            sheet_name = "Scale_Z_Number" if "Scale_Z_Number" in xls.sheet_names else xls.sheet_names[1]
            df_z = pd.read_excel(xls, sheet_name).fillna("")
            config.z_scales_a = []
            config.z_scales_b = []
            for _, row in df_z.iterrows():
                # 3 cột đầu: A
                config.z_scales_a.append(ScaleItem(
                    linguistic=str(row.iloc[0]).strip(),
                    code=str(row.iloc[1]).strip(),
                    value=row.iloc[2]
                ))
                # 3 cột sau: B (nếu có đủ 6 cột)
                if len(row) >= 6:
                    config.z_scales_b.append(ScaleItem(
                        linguistic=str(row.iloc[3]).strip(),
                        code=str(row.iloc[4]).strip(),
                        value=row.iloc[5]
                    ))

    # 3. Đọc Sheet Experts (nếu chế độ Account)
    if config.expert_auth_mode == "Account" and "Experts" in xls.sheet_names:
        df_exp = pd.read_excel(xls, "Experts").fillna("")
        config.experts = []
        for _, row in df_exp.iterrows():
            config.experts.append(Expert(
                name=str(row.iloc[0]).strip(),
                username=str(row.iloc[1]).strip(),
                password=str(row.iloc[2]).strip()
            ))

    return config

def export_all_experts_results(file_code_path: str, file_value_path: str,
                               config: ProjectConfig,
                               all_experts_data: dict,
                               row_labels: list, col_labels: list):
    """Xuất TOÀN BỘ kết quả của tất cả chuyên gia vào 2 file Excel riêng biệt:
       Mỗi Chuyên gia là 1 Sheet."""
    
    # 1. Xuất file 1: ALL_EXPERTS_RAW_CODES
    with pd.ExcelWriter(file_code_path, engine="openpyxl") as writer_code:
        # Sheet tổng quan Metadata
        meta_list = [
            {"Parameter": "Model_Type", "Value": config.model_type},
            {"Parameter": "Use_Z_Number", "Value": "YES" if config.use_z_number else "NO"},
            {"Parameter": "Scale_Type", "Value": config.scale_type},
            {"Parameter": "Total_Experts_Saved", "Value": len(all_experts_data)}
        ]
        pd.DataFrame(meta_list).to_excel(writer_code, sheet_name="General_Info", index=False)

        for exp_name, data in all_experts_data.items():
            # Tên sheet Excel tối đa 31 ký tự, loại bỏ ký tự cấm
            safe_sheet_name = exp_name.replace(":", "_").replace("/", "_")[:30]
            df = pd.DataFrame(data["code_matrix"], index=row_labels, columns=col_labels)
            df.to_excel(writer_code, sheet_name=safe_sheet_name)

    # 2. Xuất file 2: ALL_EXPERTS_NUMERICAL_VALUES
    with pd.ExcelWriter(file_value_path, engine="openpyxl") as writer_val:
        pd.DataFrame(meta_list).to_excel(writer_val, sheet_name="General_Info", index=False)

        for exp_name, data in all_experts_data.items():
            safe_sheet_name = exp_name.replace(":", "_").replace("/", "_")[:30]
            df = pd.DataFrame(data["value_matrix"], index=row_labels, columns=col_labels)
            df.to_excel(writer_val, sheet_name=safe_sheet_name)

if __name__ == "__main__":
    generate_default_template()