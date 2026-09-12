如何在本地构建 PPTX（说明）

先决条件
- 在本地安装 Python 3.8+
- 安装依赖：pip install -r deliverables/requirements.txt

步骤
1. 克隆仓库：
   git clone https://github.com/jm6vfc765k-hue/fuyun-ai.git
2. 进入项目目录：
   cd fuyun-ai
3. 安装依赖：
   pip install -r deliverables/requirements.txt
4. 运行构建脚本：
   python deliverables/build_pptx.py

输出
- deliverables/front_half_draft_v1.pptx
- deliverables/charts/*.png （已由脚本从 SVG 转换）

说明
- 如果你需要我继续在此环境尝试直接生成并上传二进制 PPTX，我可以继续尝试；但当前我已将所有构建脚本、SVG 与讲稿放到仓库，按上述步骤在本地即可还原最终 PPTX。
