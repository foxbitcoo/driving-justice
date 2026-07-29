import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SKILL_TEXT = (PROJECT_ROOT / "SKILL.md").read_text(encoding="utf-8")
TRAFFIC_LAW_TEXT = (PROJECT_ROOT / "references" / "traffic-law.md").read_text(
    encoding="utf-8"
)
PLATE_VERIFICATION_TEXT = (
    PROJECT_ROOT / "references" / "plate-verification.md"
).read_text(encoding="utf-8")


class SkillScopeTests(unittest.TestCase):
    def test_multimodal_capability_gate_runs_before_media_analysis(self):
        capability_reference_path = (
            PROJECT_ROOT / "references" / "model-capability.md"
        )

        self.assertTrue(capability_reference_path.is_file())
        capability_text = capability_reference_path.read_text(encoding="utf-8")
        self.assertIn("references/model-capability.md", SKILL_TEXT)
        self.assertLess(
            SKILL_TEXT.index("模型多模态能力预检"),
            SKILL_TEXT.index("### Step 0：前置检查"),
        )
        self.assertIn("图像输入能力", capability_text)
        self.assertIn("原生视频理解", capability_text)
        self.assertIn("本地抽帧", capability_text)
        self.assertIn("能力未知", capability_text)
        self.assertIn("停止", capability_text)
        self.assertIn("⚠️", capability_text)
        self.assertIn("不要仅凭模型名称", capability_text)

    def test_one_clear_plate_side_with_continuous_trajectory_is_accepted(self):
        self.assertIn("前牌或后牌", SKILL_TEXT)
        self.assertIn("前牌或后牌", PLATE_VERIFICATION_TEXT)
        self.assertIn("至少一面", PLATE_VERIFICATION_TEXT)
        self.assertNotIn("前牌、后牌、连续轨迹三方交叉验证", SKILL_TEXT)
        self.assertNotIn("三者缺一不可", PLATE_VERIFICATION_TEXT)

    def test_truck_using_small_passenger_only_lane_is_supported(self):
        supported_behavior = "货车违反车道通行规定"

        self.assertIn(supported_behavior, SKILL_TEXT)
        self.assertIn(f"| {supported_behavior}", TRAFFIC_LAW_TEXT)
        self.assertIn("《道路交通安全法》第三十七条、第三十八条", TRAFFIC_LAW_TEXT)


if __name__ == "__main__":
    unittest.main()
