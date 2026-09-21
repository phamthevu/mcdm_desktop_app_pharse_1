from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional

@dataclass
class Factor:
    type: str          # 'Factor' hoặc 'Alternative'
    code: str          # C1, C2, A1...
    name: str          # Tên chi tiết
    description: str = ""

@dataclass
class ScaleItem:
    linguistic: str    # VD: Very High
    code: str          # VD: VH
    value: Any         # VD: 0.75, 1.0, 1.0 hoặc 5 hoặc '0.75,1.0,1.0'

@dataclass
class Expert:
    name: str
    username: str
    password: str

@dataclass
class ProjectConfig:
    # Cấu hình trên UI
    model_type: str = "DEMATEL"       # DEMATEL, TOPSIS, AHP, DELPHI
    use_z_number: bool = False        # True / False
    scale_type: str = "Category"      # 'Category' hoặc 'Continuous'
    continuous_min: float = 0.0
    continuous_max: float = 10.0
    expert_auth_mode: str = "Account" # 'Account' (dùng file) hoặc 'Free' (tự do)
    
    # Dữ liệu đọc từ file Excel
    excel_path: str = ""
    factors: List[Factor] = field(default_factory=list)
    alternatives: List[Factor] = field(default_factory=list)
    
    # Thang đo
    # Khi ko có Z: dùng standard_scales
    standard_scales: List[ScaleItem] = field(default_factory=list)
    # Khi có Z: gồm thang đánh giá A và thang chắc chắn B
    z_scales_a: List[ScaleItem] = field(default_factory=list)
    z_scales_b: List[ScaleItem] = field(default_factory=list)
    
    # Chuyên gia
    experts: List[Expert] = field(default_factory=list)