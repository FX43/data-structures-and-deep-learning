# -*- coding: utf-8 -*-
r"""
将 D:\Data Structure\leetcode 目录下各 .ipynb 刷题笔记整理为
《LeetCode Hot100与高频题目刷题清单与进度追踪表.xlsx》
频次 / 状态 / 刷题次数(第1~3轮) 仅保留表头,内容留空待填。
"""
import json
import math
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. 题目数据:按分类分组,顺序即分类内排名
#    每项: (lc_id, 题目名, 难度, slug, 来源笔记, 标题cell序号)
# ---------------------------------------------------------------------------
CATEGORIES = [
    ("滑动窗口 + 双指针", [
        (167, "两数之和 II - 输入有序数组", "中等", "two-sum-ii-input-array-is-sorted", "双指针问题.ipynb", 0),
        (15,  "三数之和", "中等", "3sum", "双指针问题.ipynb", 2),
        (42,  "接雨水", "困难", "trapping-rain-water", "双指针问题.ipynb", 4),
        (209, "长度最小的子数组", "中等", "minimum-size-subarray-sum", "双指针问题.ipynb", 10),
        (713, "乘积小于 K 的子数组", "中等", "subarray-product-less-than-k", "双指针问题.ipynb", 12),
        (3,   "无重复字符的最长子串", "中等", "longest-substring-without-repeating-characters", "双指针问题.ipynb", 14),
        (76,  "最小覆盖子串", "困难", "minimum-window-substring", "双指针问题.ipynb", 17),
        (438, "找到字符串中所有字母异位词", "中等", "find-all-anagrams-in-a-string", "双指针问题.ipynb", 19),
    ]),
    ("二分查找", [
        (33,  "搜索旋转排序数组", "中等", "search-in-rotated-sorted-array", "二分查找错题集.ipynb", 5),
        (153, "寻找旋转排序数组中的最小值", "中等", "find-minimum-in-rotated-sorted-array", "二分查找错题集.ipynb", 7),
        (34,  "在排序数组中查找元素的第一个和最后一个位置", "中等", "find-first-and-last-position-of-element-in-sorted-array", "二分查找错题集.ipynb", 9),
        (162, "寻找峰值", "中等", "find-peak-element", "二分查找错题集.ipynb", 11),
        (75,  "颜色分类", "中等", "sort-colors", "二分查找错题集.ipynb", 13),
        (4,   "寻找两个正序数组的中位数", "困难", "median-of-two-sorted-arrays", "二分查找错题集.ipynb", 15),
    ]),
    ("动态规划", [
        (198,  "打家劫舍", "中等", "house-robber", "动态规划问题.ipynb", 2),
        (152,  "乘积最大子数组", "中等", "maximum-product-subarray", "动态规划问题.ipynb", 4),
        (494,  "目标和", "中等", "target-sum", "动态规划问题.ipynb", 6),
        (322,  "零钱兑换", "中等", "coin-change", "动态规划问题.ipynb", 12),
        (1143, "最长公共子序列", "中等", "longest-common-subsequence", "动态规划问题.ipynb", 17),
        (712,  "两个字符串的最小ASCII删除和", "中等", "minimum-ascii-delete-sum-for-two-strings", "动态规划问题.ipynb", None),
        (72,   "编辑距离", "中等", "edit-distance", "动态规划问题.ipynb", 23),
        (300,  "最长递增子序列", "中等", "longest-increasing-subsequence", "动态规划问题.ipynb", 26),
        (1964, "找出到每个位置为止的最长合法障碍赛跑路线", "困难", "longest-obstacle-course-at-each-position", "动态规划问题.ipynb", 31),
        (1671, "得到山形数组的最少删除次数", "困难", "minimum-number-of-removals-to-make-mountain-array", "动态规划问题.ipynb", None),
        (122,  "买卖股票的最佳时机 II", "中等", "best-time-to-buy-and-sell-stock-ii", "动态规划问题.ipynb", 34),
        (309,  "买卖股票的最佳时机含冷冻期", "中等", "best-time-to-buy-and-sell-stock-with-cooldown", "动态规划问题.ipynb", 38),
        (188,  "买卖股票的最佳时机 IV", "困难", "best-time-to-buy-and-sell-stock-iv", "动态规划问题.ipynb", 40),
        (516,  "最长回文子序列", "中等", "longest-palindromic-subsequence", "动态规划问题.ipynb", 46),
        (1039, "多边形三角剖分的最低得分", "中等", "minimum-score-triangulation-of-polygon", "动态规划问题.ipynb", 48),
        (647,  "回文子串", "中等", "palindromic-substrings", "动态规划问题.ipynb", 50),
        (5,    "最长回文子串", "中等", "longest-palindromic-substring", "动态规划问题.ipynb", 52),
        (32,   "最长有效括号", "困难", "longest-valid-parentheses", "动态规划问题.ipynb", 54),
        (139,  "单词拆分", "中等", "word-break", "动态规划问题.ipynb", 56),
        (1187, "使数组严格递增", "困难", "make-array-strictly-increasing", "动态规划问题.ipynb", 58),
        (2266, "统计打字方案数", "中等", "count-number-of-texts", "动态规划问题.ipynb", 61),
        (1770, "执行乘法运算的最大分数", "困难", "maximum-score-from-performing-multiplication-operations", "动态规划问题.ipynb", 63),
        (2246, "相邻字符不同的最长路径", "困难", "longest-path-with-different-adjacent-characters", "动态规划问题.ipynb", 65),
        (2538, "最大价值和与最小价值和的差值", "困难", "difference-between-maximum-and-minimum-price-sum", "动态规划问题.ipynb", 67),
        (1617, "统计子树中城市之间的最大距离", "困难", "count-subtrees-with-max-distance-between-cities", "动态规划问题.ipynb", 69),
        (337,  "打家劫舍 III", "中等", "house-robber-iii", "动态规划问题.ipynb", 76),
        (968,  "监控二叉树", "困难", "binary-tree-cameras", "动态规划问题.ipynb", 79),
    ]),
    ("贪心算法", [
        (55, "跳跃游戏", "中等", "jump-game", "贪心算法.ipynb", 0),
        (45, "跳跃游戏 II", "中等", "jump-game-ii", "贪心算法.ipynb", 6),
    ]),
    ("回溯", [
        (46,  "全排列", "中等", "permutations", "回溯问题.ipynb", 1),
        (51,  "N 皇后", "困难", "n-queens", "回溯问题.ipynb", 5),
        (17,  "电话号码的字母组合", "中等", "letter-combinations-of-a-phone-number", "回溯问题.ipynb", 9),
        (78,  "子集", "中等", "subsets", "回溯问题.ipynb", 11),
        (131, "分割回文串", "中等", "palindrome-partitioning", "回溯问题.ipynb", 16),
        (39,  "组合总和", "中等", "combination-sum", "回溯问题.ipynb", 18),
        (22,  "括号生成", "中等", "generate-parentheses", "回溯问题.ipynb", 20),
        (216, "组合总和 III", "中等", "combination-sum-iii", "回溯问题.ipynb", 22),
    ]),
    ("链表", [
        (206, "反转链表", "简单", "reverse-linked-list", "链表错题集.ipynb", 7),
        (160, "相交链表", "简单", "intersection-of-two-linked-lists", "链表错题集.ipynb", 8),
        (141, "环形链表", "简单", "linked-list-cycle", "链表错题集.ipynb", 10),
        (142, "环形链表 II", "中等", "linked-list-cycle-ii", "链表错题集.ipynb", 12),
        (234, "回文链表", "简单", "palindrome-linked-list", "链表错题集.ipynb", 14),
        (24,  "两两交换链表中的节点", "中等", "swap-nodes-in-pairs", "链表错题集.ipynb", 16),
        (148, "排序链表", "中等", "sort-list", "链表错题集.ipynb", 17),
        (23,  "合并 K 个升序链表", "困难", "merge-k-sorted-lists", "链表错题集.ipynb", 19),
        (25,  "K 个一组翻转链表", "困难", "reverse-nodes-in-k-group", "链表错题集.ipynb", 21),
        (138, "随机链表的复制", "中等", "copy-list-with-random-pointer", "链表错题集.ipynb", 23),
    ]),
    ("栈 / 单调栈", [
        (394, "字符串解码", "中等", "decode-string", "栈错题集.ipynb", 0),
        (155, "最小栈", "中等", "min-stack", "栈错题集.ipynb", 2),
        (739, "每日温度", "中等", "daily-temperatures", "栈错题集.ipynb", 4),
        (84,  "柱状图中最大的矩形", "困难", "largest-rectangle-in-histogram", "栈错题集.ipynb", 9),
    ]),
    ("队列 / 单调队列", [
        (239, "滑动窗口最大值", "困难", "sliding-window-maximum", "单调队列问题.ipynb", 0),
    ]),
    ("哈希表", [
        (128, "最长连续序列", "中等", "longest-consecutive-sequence", "哈希表错题集.ipynb", 0),
        (49,  "字母异位词分组", "中等", "group-anagrams", "链表错题集.ipynb", 0),
    ]),
    ("堆", [
        (295, "数据流的中位数", "困难", "find-median-from-data-stream", "堆错题集.ipynb", 0),
    ]),
    ("树", [
        (101, "对称二叉树", "简单", "symmetric-tree", "树错题集.ipynb", 1),
        (543, "二叉树的直径", "简单", "diameter-of-binary-tree", "树错题集.ipynb", 3),
        (104, "二叉树的最大深度", "简单", "maximum-depth-of-binary-tree", "树错题集.ipynb", 5),
        (226, "翻转二叉树", "简单", "invert-binary-tree", "树错题集.ipynb", 7),
        (102, "二叉树的层序遍历", "中等", "binary-tree-level-order-traversal", "树错题集.ipynb", 9),
        (105, "从前序与中序遍历序列构造二叉树", "中等", "construct-binary-tree-from-preorder-and-inorder-traversal", "树错题集.ipynb", 11),
        (98,  "验证二叉搜索树", "中等", "validate-binary-search-tree", "树错题集.ipynb", 18),
        (114, "二叉树展开为链表", "中等", "flatten-binary-tree-to-linked-list", "树错题集.ipynb", 20),
        (236, "二叉树的最近公共祖先", "中等", "lowest-common-ancestor-of-a-binary-tree", "树错题集.ipynb", 23),
        (124, "二叉树中的最大路径和", "困难", "binary-tree-maximum-path-sum", "树错题集.ipynb", 25),
        (437, "路径总和 III", "中等", "path-sum-iii", "树错题集.ipynb", 27),
    ]),
    ("图", [
        (200, "岛屿数量", "中等", "number-of-islands", "图问题.ipynb", 0),
        (207, "课程表", "中等", "course-schedule", "图问题.ipynb", 5),
        (208, "实现 Trie (前缀树)", "中等", "implement-trie-prefix-tree", "图问题.ipynb", 8),
    ]),
    ("矩阵", [
        (54, "螺旋矩阵", "中等", "spiral-matrix", "矩阵问题.ipynb", 1),
        (48, "旋转图像", "中等", "rotate-image", "矩阵问题.ipynb", 3),
        (74, "搜索二维矩阵", "中等", "search-a-2d-matrix", "矩阵问题.ipynb", 5),
        (73, "矩阵置零", "中等", "set-matrix-zeroes", "矩阵问题.ipynb", 7),
    ]),
    ("数组", [
        (53,  "最大子数组和", "中等", "maximum-subarray", "数组错题集.ipynb", 0),
        (238, "除自身以外数组的乘积", "中等", "product-of-array-except-self", "数组错题集.ipynb", 2),
        (560, "和为 K 的子数组", "中等", "subarray-sum-equals-k", "数组错题集.ipynb", 4),
        (189, "轮转数组", "中等", "rotate-array", "数组错题集.ipynb", 6),
        (56,  "合并区间", "中等", "merge-intervals", "数组错题集.ipynb", 10),
        (215, "数组中的第 K 个最大元素", "中等", "kth-largest-element-in-an-array", "数组错题集.ipynb", 12),
        (287, "寻找重复数", "中等", "find-the-duplicate-number", "数组错题集.ipynb", 14),
        (41,  "缺失的第一个正数", "困难", "first-missing-positive", "数组错题集.ipynb", 16),
        (31,  "下一个排列", "中等", "next-permutation", "数组错题集.ipynb", 19),
        (283, "移动零", "简单", "move-zeroes", "链表错题集.ipynb", 4),
    ]),
]

# ---------------------------------------------------------------------------
# 2. 解题思路(手写精简版);未覆盖的题目从笔记 markdown 提取
# ---------------------------------------------------------------------------
NOTE = {
    167: "相向双指针：数组有序，和<target左指针右移，和>target右指针左移（题目下标从1开始）",
    15:  "固定第一个数x，剩余区间相向双指针；i、j、k 三处都要跳过重复元素去重",
    42:  "三种解法：前后缀最大值 / 相向双指针 / 单调栈；每格存水量=min(左最高,右最高)-height",
    76:  "滑动窗口：先扩右指针直到覆盖t，再收缩左指针求最小长度；用need/剩余字符计数判断覆盖",
    206: "迭代：prev/curr指针逐个反转next；递归：先反转后续，再 head.next.next=head",
    24:  "需要dummy节点；pre指向每组前驱，两两交换a、b后移动指针",
    51:  "回溯逐行放皇后；用列、主对角线(i-j)、副对角线(i+j)集合剪枝",
    22:  "回溯选与不选：左括号数<n才可加'('，右括号数<左括号数才可加')'",
    216: "回溯组合：从1~9选k个数和为n；可剪枝（剩余个数不足或和超target）",
    155: "辅助栈同步压入当前最小值；或每个元素存(值, 当前最小值)",
    101: "递归：两棵子树镜像对称（左.左==右.右 且 左.右==右.左）；或队列逐对比较",
    543: "递归求每个节点左右子树深度，全局max(左深+右深)；直径不一定要过根",
    104: "递归：max(左子树深度, 右子树深度)+1；空节点返回0",
    226: "递归交换左右孩子；或层序遍历逐层交换",
    98:  "中序遍历结果严格递增；或递归传递(lower, upper)上下界校验",
    236: "递归：节点为p或q返回自身；左右子树返回值均非空时当前节点即最近公共祖先",
    124: "递归返回单边最大路径和 node.val+max(左,右,0)；全局max记录经过当前节点的最大路径和",
    54:  "按上、右、下、左四条边界循环收缩；防止m*n为奇数时中心元素遗漏",
    48:  "先转置（沿主对角线），再逐行反转",
    74:  "每行有序→每行二分；或从右上角出发，小于target行+1，大于target列-1",
    73:  "用首行首列作为标记数组记录含0的行列，再二次遍历置零；注意首行首列本身要先处理",
    55:  "贪心：维护能到达的最远位置reach，遍历时更新；reach>=n-1即成功",
    208: "Trie节点含children字典（或26数组）与isEnd标记；插入/搜索/前缀搜索三步走",
    33:  "红蓝染色/二分：先判断mid落在左段还是右段有序区间，再决定搜索方向",
    162: "红蓝染色：nums[mid]<nums[mid+1] 说明峰在右侧，否则峰在左侧",
    75:  "荷兰国旗三指针：[0,zero]全0、[two,n-1]全2、中间为1；partition交换",
    4:   "二分排除：在较短数组上二分划分，保证左半最大<右半最小；注意奇偶长度中位数处理",
    215: "快速选择partition（随机pivot）期望O(n)；或维护大小为k的小顶堆",
    122: "状态机DP：dp[i][0]不持有、dp[i][1]持有；不限制交易次数，答案=dp[n-1][0]",
    309: "含冷冻期：买入只能从两天前(f2)转移；f0卖出/f1持有/f2冷冻 滚动更新",
    188: "至多k笔：dp[i][j][0/1]，j为已交易次数；j<0不合法，需k+2维空间",
    337: "树形DP：每个节点返回[偷,不偷]；偷则子节点不能偷，不偷取子节点最大值；后序遍历",
    2246: "树形DP：求经过每个节点的最长路径，仅相邻字符不同才可连边；维护第一大/第二大分支",
    2538: "树形DP：维护每个节点向下的路径价值最大和与最小和；答案为max-min，可在节点处拼接两条路径",
    32:  "栈存左括号下标，栈底保留最后一个未匹配的')'作哨兵；遇')'弹出并更新ans",
    516: "区间DP：f[i][j]表示s[i..j]的最长回文子序列；s[i]==s[j]则+2，否则取max(去左,去右)",
    712: "LCS变式：总ASCII和-2×LCS；或直接DP，两字符不同时取删s1[i]或删s2[j]的较小值",
    1671: "山形数组：LIS+LDS，枚举峰顶，删除数=n-max(LIS[i]+LDS[i]-1)；峰两侧都需有元素",
    494:  "01背包：target += sum(nums)，求恰好等于target的方案数（每个数选或不选）",
    322: "完全背包：f[i]=min(f[i], f[i-coin]+1)；f初始化为inf、f[0]=0，返回时判断是否可达",
    1143: "二维DP：text1[i]==text2[j]则f[i][j]=f[i-1][j-1]+1，否则取max(f[i-1][j], f[i][j-1])",
    72:  "DP：f[i][j]编辑距离；末尾字符相同直接继承，不同则取增/删/改三者最小值+1",
    2266: "组合计数+DP：连续相同数字分组，每组按键方案满足dp[i]=dp[i-1]+dp[i-2]+dp[i-3]+(4字母键再加dp[i-4])；结果取模1e9+7",
}

# ---------------------------------------------------------------------------
# 3. 从笔记提取 markdown 备注
# ---------------------------------------------------------------------------
def clean_md(text):
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", text)          # 图片
    text = text.replace("**_", "").replace("**", "")
    text = re.sub(r"[*_#`>]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def load_notes():
    """{cell_idx: 该标题cell的备注文本} 按笔记文件"""
    result = {}
    titles = {}  # (nb, idx) -> 题目名(去空格)
    for cat, items in CATEGORIES:
        for _id, _t, _d, _slug, nb, idx in items:
            if idx is None:
                continue
            result.setdefault(nb, {})
            titles[(nb, idx)] = _t.replace(" ", "")
    for nb in result:
        with open(nb, "r", encoding="utf-8") as f:
            cells = json.load(f)["cells"]
        used = sorted({idx for cat, items in CATEGORIES for _id, _t, _d, _slug, n, idx
                       in items if n == nb and idx is not None})
        for h in used:
            end = used[used.index(h) + 1] if used.index(h) + 1 < len(used) else len(cells)
            raw = "".join(cells[h]["source"])
            parts = []
            pending = ""  # 无 ** 标题里可能带正文,仅当后续无备注时使用
            if "**" in raw:
                # 取第二个 ** 之后的正文(很多笔记把思路写在标题同一格)
                after = clean_md(raw.split("**", 2)[-1]).strip("：: ")
                if after:
                    parts.append(after)
            else:
                rest = clean_md(re.sub(r"^#\s*\d*\.?\s*", "", raw))
                t = titles.get((nb, h), "")
                if t and rest.lower().startswith(t.lower()):
                    rest = rest[len(t):].strip("：: ")
                if len(rest) > 8:
                    pending = rest
            for i in range(h + 1, end):
                if cells[i]["cell_type"] != "markdown":
                    continue
                txt = clean_md("".join(cells[i]["source"]))
                if not txt:
                    continue
                # 跳过被误贴进 markdown 单元格的代码
                if re.match(r"^(class |def |import |from )", txt) \
                        or txt.startswith("class Solution"):
                    continue
                if txt not in parts:
                    parts.append(txt)
            if not parts and pending:
                parts.append(pending)
            result[nb][h] = " ".join(parts)
    return result

def truncate(text, limit=110):
    return text if len(text) <= limit else text[:limit].rstrip("，。；、 ") + "…"

NOTES = load_notes()

def note_for(_id, nb, idx):
    if _id in NOTE:
        return NOTE[_id]
    if idx is None or idx not in NOTES.get(nb, {}):
        return ""
    return truncate(NOTES[nb][idx])

# ---------------------------------------------------------------------------
# 4. 汇总区与表头
# ---------------------------------------------------------------------------
rows = []  # (分类, 排名, 编号, 题目, 难度, 频次, 状态, 链接, 刷题次数, 备注)
for cat, items in CATEGORIES:
    for rank, (_id, title, diff, slug, nb, idx) in enumerate(items, 1):
        rows.append([
            cat, rank, _id, f"{_id}. {title}", diff, "", "",
            f"https://leetcode.cn/problems/{slug}/", "",
            note_for(_id, nb, idx),
        ])

TOTAL = len(rows)

wb = Workbook()
ws = wb.active
ws.title = "Hot100与高频题目"

thin = Side(style="thin", color="B0B0B0")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

# 列宽
widths = {"A": 17, "B": 9, "C": 12, "D": 34, "E": 9, "F": 8,
          "G": 11, "H": 42, "I": 14, "J": 60}
for col, w in widths.items():
    ws.column_dimensions[col].width = w

NCOL = 10  # A..J

# --- 汇总区(第1-2行) ---
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOL)
c = ws.cell(row=1, column=1, value=f"LeetCode Hot 100 与高频题目总数：{TOTAL}")
c.font = Font(name="微软雅黑", size=16, bold=True, color="FFFFFF")
c.fill = PatternFill("solid", fgColor="1F4E79")
c.alignment = center
ws.row_dimensions[1].height = 34

ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NCOL)
c = ws.cell(row=2, column=1,
            value=("说明：本表由 D:\\Data Structure\\leetcode 目录下 14 个 .ipynb 刷题笔记整理生成，"
                   f"共 {TOTAL} 题，按算法分类分组、分类内按笔记出现顺序排名。"
                   "频次、状态暂仅保留表头待填写；刷题次数每行留一个空格，可直接填写 ✓ 打卡记录复习情况；"
                   "难度与链接为 LeetCode 官方信息，解题思路来自笔记原文提炼。"))
c.font = Font(name="微软雅黑", size=10, color="595959")
c.fill = PatternFill("solid", fgColor="DDEBF7")
c.alignment = left_wrap
ws.row_dimensions[2].height = 46

# --- 表头(第3行,单行) ---
headers = ["分类 (Category)", "分类排名 (Category Rank)", "LeetCode编号 (LeetCode ID)",
           "题目 (Problem Title)", "难度 (Difficulty)", "频次 (Frequency)",
           "状态 (Status Tag)", "LeetCode链接 (URL)", "刷题次数 (Practice Rounds)",
           "解题思路与注意事项 (Notes/Key Takeaways)"]
hdr_fill = PatternFill("solid", fgColor="1F4E79")
hdr_font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
for col_idx, h in enumerate(headers, 1):
    cell = ws.cell(row=3, column=col_idx, value=h)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = center
    cell.border = border
ws.row_dimensions[3].height = 30

# --- 数据区(第4行起) ---
cat_fills = {
    "滑动窗口 + 双指针": "DDEBF7", "二分查找": "E2EFDA", "动态规划": "FFF2CC",
    "贪心算法": "FCE4D6", "回溯": "E6E0EC", "链表": "D9EAD3",
    "栈 / 单调栈": "F4CCCC", "队列 / 单调队列": "CFE2F3", "哈希表": "FFE599",
    "堆": "C9DAF8", "树": "D0E0E3", "图": "EAD1DC", "矩阵": "D9D2E9",
    "数组": "F9CB9C",
}
diff_font = {"简单": "1E7B34", "中等": "B45F06", "困难": "C00000"}
diff_fill = {"简单": "E2EFDA", "中等": "FFF2CC", "困难": "F4CCCC"}

first_data = 4
for i, r in enumerate(rows):
    rn = first_data + i
    cat, rank, lc_id, title, diff, freq, status, url, rounds, note = r
    ws.cell(row=rn, column=1, value=cat).font = Font(name="微软雅黑", size=10, bold=True)
    ws.cell(row=rn, column=2, value=rank)
    ws.cell(row=rn, column=3, value=lc_id)
    ws.cell(row=rn, column=4, value=title)
    ws.cell(row=rn, column=5, value=diff)
    ws.cell(row=rn, column=8, value=url)
    ws.cell(row=rn, column=10, value=note)

    for col_idx in range(1, NCOL + 1):
        cell = ws.cell(row=rn, column=col_idx)
        cell.border = border
        if col_idx in (6, 7):                      # 频次 / 状态:占位灰底
            cell.fill = PatternFill("solid", fgColor="F2F2F2")
            cell.alignment = center
        elif col_idx == 1:
            cell.fill = PatternFill("solid", fgColor=cat_fills[cat])
            cell.alignment = center
        elif col_idx in (2, 3):
            cell.alignment = center
        elif col_idx == 4:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.font = Font(name="微软雅黑", size=10)
        elif col_idx == 5:
            cell.font = Font(name="微软雅黑", size=10, bold=True, color=diff_font[diff])
            cell.fill = PatternFill("solid", fgColor=diff_fill[diff])
            cell.alignment = center
        elif col_idx == 8:
            # 统一显示:去掉每格重复的 https:// 前缀,固定列宽+自动换行,长链接转第二行
            display = url.replace("https://", "")
            cell.value = display
            cell.hyperlink = url
            cell.font = Font(name="微软雅黑", size=9, color="0563C1", underline="single")
            cell.alignment = left_wrap
        elif col_idx == 9:                         # 刷题次数:每行留一个空格,自行填 ✓
            cell.alignment = center
            cell.font = Font(name="微软雅黑", size=12)
        elif col_idx == 10:
            cell.alignment = left_wrap
            cell.font = Font(name="微软雅黑", size=9)
    # 行高:同时容纳备注与链接(链接最多2行)所需行数,避免截断或溢出
    est_note_lines = max(1, math.ceil(len(note) / 26))
    est_url_lines = max(1, math.ceil(len(url.replace("https://", "")) / 40))
    ws.row_dimensions[rn].height = max(24, max(est_note_lines, est_url_lines) * 14 + 8)

# --- 分类合并(第1列) ---
start = first_data
for cat, items in CATEGORIES:
    end = start + len(items) - 1
    if end > start:
        ws.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)
    start = end + 1

# 冻结窗格:冻结前3行与前2列
ws.freeze_panes = "C4"

out = "LeetCode Hot100与高频题目刷题清单与进度追踪表.xlsx"
wb.save(out)
print(f"已生成 {out}  |  题目总数: {TOTAL}  |  分类数: {len(CATEGORIES)}")
for cat, items in CATEGORIES:
    print(f"  {cat}: {len(items)} 题")
