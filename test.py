# 调用方法前需先对nums数组排序
import copy


def nsum_target(nums, n, start, target):
    sz = len(nums)
    res = []
    # 至少是 2Sum，且数组大小不应该小于 n
    if n < 2 or sz < n:
        return res
    # 2Sum 是 base case
    if n == 2:
        # 双指针那一套操作
        lo = start
        hi = sz - 1
        while lo < hi:
            sum_ = nums[lo] + nums[hi]
            left, right = nums[lo], nums[hi]
            if sum_ < target:
                while lo < hi and nums[lo] == left:
                    lo += 1
            elif sum_ > target:
                while lo < hi and nums[hi] == right:
                    hi -= 1
            else:
                res.append([left, right])
                while lo < hi and nums[lo] == left:
                    lo += 1
                while lo < hi and nums[hi] == right:
                    hi -= 1
    else:
        # n > 2 时，递归计算 (n-1)Sum 的结果
        for i in range(start, sz):
            if i < sz - 1 and nums[i] == nums[i + 1]:
                continue
            sub = nsum_target(nums, n - 1, i + 1, target - nums[i])
            for arr in sub:
                # (n-1)Sum 加上 nums[i] 就是 nSum
                arr.append(nums[i])
                res.append(arr)
    return res


# Definition for a binary tree node.
class TreeNode:
    def __init__(self, x):
        self.val = x
        self.left = None
        self.right = None


class LinkedList:
    def __init__(self, x):
        self.val = x
        self.next = None


def sorted_array_to_BST(nums):
    res = []

    def helper(start, end):
        if start > end:
            return None
        mid = (end + start) // 2
        node = nums[mid]
        res.append(node)
        ans = TreeNode(node)
        ans.left = helper(start, mid - 1)
        ans.right = helper(mid + 1, end)
        return ans, res

    return helper(0, len(nums) - 1)


def find_mid(head, tail):
    fast, low = head, head
    while fast.next != tail and fast != tail:
        fast, low = fast.next.next, low.next
    return low


def buildTree(head, tail):
    if head == tail:
        return None
    mid = find_mid(head, tail)
    node = mid.val
    res = TreeNode(node)
    res.left = buildTree(head, mid)
    res.right = buildTree(mid.next, tail)
    return res


def Btree_to_linkedlist(root):
    # 二叉树转链表
    while not root.left:
        if not root.left:
            root = root.right
        else:
            pre = root.left
            while pre.right:
                pre = pre.right
            pre.right = root.right
            root.right = root.left
            root.left = None
            root = root.right
    return root


def func_121(prices):
    # leetcode121题，股票一次买卖
    if len(prices) <= 1:
        return 0
    dp = [0 for i in prices]
    min_price = prices[0]
    for i in range(1, len(prices)):
        min_price = min(min_price, prices[i])
        dp[i] = max(dp[i - 1], prices[i] - min_price)
    return dp


def func_122(prices):
    # 122题，股票多次买卖
    if len(prices) <= 1:
        return 0
    res = 0
    for i in range(1, len(prices)):
        if prices[i] > prices[i - 1]:
            res += prices[i] - prices[i - 1]
    return res


def tree_preorder(root):
    # 先序遍历
    res = []

    def helper(root):
        if not root:
            return
        res.append(root.val)
        helper(root.left)
        helper(root.right)
        return res

    helper(root)
    return res


def tree_postorder(root):
    # 后续遍历
    res = []

    def helper(root):
        if not root:
            return
        helper(root.left)
        helper(root.right)
        res.append(root.val)
        return res

    helper(root)
    return res


def max_huiwen(string):
    sz = len(string)
    if sz <= 1: return string
    dp = [[False] * sz] * sz
    max_v = 1
    start = 0
    for i in range(sz):
        dp[i][i] = True
    for right in range(1, sz):
        for left in range(right):
            if string[right] == string[left]:
                if right - left < 3:
                    dp[left][right] = True
                else:
                    dp[left][right] = dp[left + 1][right - 1]
            else:
                dp[left][right] = False
            if dp[left][right]:
                cur = right - left + 1
                if cur > max_v:
                    max_v = max(right - left + 1, max_v)
                    start = left
    return string[start: start + max_v]


def min_trace_sum_64(grid):
    # grid: 2d,[[1,3,1],
    #           [1,5,1],]
    #           [4,2,1],]
    # every step can only move to up or right
    # dynamic program
    if not grid or len(grid) == 0:
        return 0
    row, col = len(grid), len(grid[0])
    dp = [[0 for _ in range(col)] for _ in range(row)]
    dp[0][0] = grid[0][0]
    # trun up:
    for r in range(1, row):
        dp[r][0] = dp[r - 1][0] + grid[r][0]
    # turn right
    for c in range(1, col):
        dp[0][c] = dp[0][c - 1] + grid[0][c]
    for i in range(1, row):
        for j in range(1, col):
            dp[i][j] = min(dp[i - 1][j], dp[i][j - 1]) + grid[i][j]
    return dp[row - 1][col - 1]


class Solution:
    def __init__(self):
        self.result = 0

    def sum_target_494(self, nums, target):
        if len(nums) == 0:
            return 0
        self.back_trace(nums, 0, target)
        return self.result

    def back_trace(self, nums, i, rest):
        # nonlocal result
        if i == len(nums):
            if rest == 0:
                self.result += 1
            return
        self.back_trace(nums, i + 1, rest - nums[i])
        self.back_trace(nums, i + 1, rest + nums[i])


def three_sum_15(nums):
    """
    :type nums: List[int]
    :rtype: List[List[int]]
    """
    nums = sorted(nums)
    res = {}
    for i, value in enumerate(nums):
        if value > 0:
            break
        target = 0 - value
        left, right = i + 1, len(nums) - 1
        while left < right:
            if nums[left] + nums[right] > target:
                right -= 1
            elif nums[left] + nums[right] < target:
                left += 1
            elif nums[left] + nums[right] == target:
                string = '{},{},{}'.format(value, nums[left], nums[right])
                if string not in res:
                    res[string] = [value, nums[left], nums[right]]
                left += 1
                right -= 1
    return list(res.values())


class Solution32:
    ## 方法三：动态规划
    def longestValidParentheses_1(self, s: str) -> int:
        # 状态：以该点结尾的最长有效括号的子串长度
        dp = [0 for _ in range(len(s))]
        if len(s) < 2:
            return 0

        if s[1] == ')' and s[0] == '(':
            dp[1] = 2

        if len(s) == 2:
            return dp[1]

        for i in range(2, len(s)):
            if s[i] == '(':  ## 一种情况
                dp[i] = 0
                continue
            if s[i] == ')':  ## 另一种情况
                if s[i - 1] == '(':  ## 细分
                    dp[i] = dp[i - 2] + 2
                if s[i - 1] == ')' and s[i - 1 - dp[i - 1]] == '(' and i - 1 - dp[i - 1] >= 0:  ## 复杂的一种情况，需要特判
                    dp[i] = dp[i - 1 - dp[i - 1] - 1] + 2 + dp[i - 1]
        print(dp)
        return max(dp)

    def longestValidParentheses_2(self, s: str) -> int:
        # 栈方法，在刚才的方法上进行优化就好了，可以减少空间复杂度
        stack = [-1]
        res = 0
        for i in range(len(s)):
            if s[i] == '(':
                stack.append(i)  ## 这种思路牛逼，不要传入什么字符串了，直接传入该括号的序号，然后一减就是长度了，牛逼～
                continue  ## 这个是用栈的最简单的方法了
            stack.pop()
            if not stack:
                stack.append(i)
            else:
                # print(i, stack)
                res = max(res, i - stack[-1])
        return res


def mergeKLists(lists):
    """
    :type lists: List[ListNode]
    :rtype: ListNode
    """
    length = len(lists)
    if not lists:
        return []
    if length == 1:
        return lists[0]
    if length == 2:
        return merge2lists(lists[0], lists[1])
    mid = length // 2
    lists_pre = lists[:mid]
    lists_post = lists[mid:]
    return merge2lists(mergeKLists(lists_pre), mergeKLists(lists_post))


def merge2lists(l1, l2):
    if not l1:
        return l2
    if not l2:
        return l1
    node = []
    if l1[0] <= l2[0]:
        node.append(l1[0])
        node.extend(merge2lists(l1[1:], l2))
    else:
        node.append(l2[0])
        node.extend(merge2lists(l1, l2[1:]))
    return node


def minimumTotal_120(triangle):
    """
    :type triangle: List[List[int]]
    :rtype: int
    dp[i][j] = max(dp[i-1][j], dp[i-1][j+1]) + triangle[i][j]
    triangle = [
                 [2],
                [3,4],
               [6,5,7],
              [4,1,8,3]
            ]
    """
    n = len(triangle)
    dp = [[0] * (n + 1) for _ in range(n + 1)]
    for i in range(len(triangle[n - 1])):
        dp[n - 1][i] = triangle[n - 1][i]
    for i in range(n - 1, -1, -1):
        for j in range(len(triangle[i])):
            dp[i][j] = min(dp[i + 1][j], dp[i + 1][j + 1]) + triangle[i][j]
    return dp[0][0]


def threeSumClosest_16(nums, target):
    """
    :type nums: List[int]
    :type target: int
    :rtype: int
    """
    sort_nums = sorted(nums)
    length = len(sort_nums)
    if length == 3:
        return sum(nums)
    if length < 3:
        return
    dic = {}
    for a, a_value in enumerate(sort_nums):
        b, c = a + 1, length - 1
        while b < c:
            sub_sum = sort_nums[b] + sort_nums[c]
            key = abs(target - a_value - sub_sum)
            if sub_sum > target - a_value:
                c -= 1
            elif sub_sum < target - a_value:
                b += 1
            else:
                return a_value + sub_sum
            dic[str(key)] = a_value + sub_sum
    dic = sorted(dic.items(), key=lambda k: int(k[0]))
    return dic[0][1]


class ListNode(object):
    def __init__(self, x):
        self.val = x
        self.next = None


def deleteDuplicates_82(head):
    """
    :type head: ListNode
    :rtype: ListNode
    """
    dummpy = pre = ListNode(0)
    cur = head
    pre.next = cur
    while cur and cur.next:
        if cur.val == cur.next.val:
            while cur and cur.next and cur.val == cur.next.val:
                cur = cur.next
            cur = cur.next
            pre.next = cur
        else:
            pre = pre.next
            cur = cur.next
    return dummpy.next


def minPathSum_64(grid):
    """
    :type grid: List[List[int]]
    :rtype: int
    """
    if not grid:
        return
    row = len(grid)
    col = len(grid[0])
    dp = grid
    for i in range(row - 1, -1, -1):
        for j in range(col):
            dp[i][j] = min(dp[i - 1][j], dp[i][j - 1]) + grid[i][j]
    return dp


def merge_88(nums1, nums2):
    """
    :type nums1: List[int]
    :type m: int
    :type nums2: List[int]
    :type n: int
    :rtype: None Do not return anything, modify nums1 in-place instead.
    """
    # s1 = sorted(nums1)
    # s2 = sorted(nums2)
    # res = []
    #
    # def helper(nums1, nums2, res):
    #     if not nums1:
    #         return res + nums2
    #     if not nums2:
    #         return res + nums1
    #     min1, min2 = nums1[0], nums2[0]
    #     if min1 > min2:
    #         res.append(min2)
    #         return helper(nums1, nums2[1:], res)
    #     else:
    #         res.append(min1)
    #         return helper(nums1[1:], nums2, res)
    #
    # return helper(s1, s2, res)
    s1 = copy.deepcopy(nums1)
    s2 = sorted(nums2)
    for j in range(len(s1)-1, -1, -1):
        if s1[j] == 0:
            s1.pop()
        else:
            break
    list.sort(s1)

    i = 0

    def helper(s1, s2, nums1, i):
        if not s1:
            for ii in s2:
                nums1[i] = ii
                i += 1
            return nums1
        if not s2:
            for ii in s1:
                nums1[i] = ii
                i += 1
            return nums1
        min1, min2 = s1[0], s2[0]
        if min1 > min2:
            nums1[i] = min2
            i += 1
            return helper(s1, s2[1:], nums1, i)
        else:
            nums1[i] = min1
            i += 1
            return helper(s1[1:], s2, nums1, i)

    return helper(s1, s2, nums1, i)


def pathSum_113(root, sum):
    """
    :type root: TreeNode
    :type sum: int
    :rtype: List[List[int]]
    """
    res = []
    temp = []

    def helper(root, sum, temp, res):
        if not root:
            return res
        temp.append(root.val)
        if not root.left and root.right:
            if root.val == sum:
                res.append(temp)
            temp = []
        return helper(root.left, sum - root.val, temp, res) or helper(root.right, sum - root.val, temp, res)

    return helper(root, sum, temp, res)


def findPeakElement_162(nums):
    """
    :type nums: List[int]
    :rtype: int
    """
    # O(logn), 二分查找
    l, r = 0, len(nums) - 1
    res = []
    while l < r:
        mid = (r + l) // 2
        if nums[mid] > nums[mid + 1]:
            res.append(mid)
            r = mid
        else:
            l = mid - 1
    return res


def permute_46(self, nums):
    """
    :type nums: List[int]
    :rtype: List[List[int]]
    """
    # 排列问题，使用used
    if not len(nums):
        return []
    res, temp = [], []

    def dfs(nums, used, temp, res):
        if len(temp) == len(nums):
            res.append(temp[:])
            return
        for i in range(len(nums)):
            if not used[i]:
                used[i] = True
                dfs(nums, used, temp + [nums[i]], res)
                used[i] = False

    used = [False for _ in range(len(nums))]
    dfs(nums, used, temp, res)
    return res


def combinationSum_39(self, candidates, target):
    """
    :type candidates: List[int]
    :type target: int
    :rtype: List[List[int]]
    """
    # 回溯算法, 递归+深度优先遍历
    # 不重复元素，则使用begin
    if not len(candidates):
        return
    res, temp = [], []
    candidates.sort()

    def dfs(candidates, begin, res, temp, target):
        if target == 0:
            res.append(temp)
            return
        for i in range(begin, len(candidates)):
            rest = target - candidates[i]
            if rest < 0:
                break
            dfs(candidates, i, res, temp + [candidates[i]], rest)

    dfs(candidates, 0, res, temp, target)
    return res


if __name__ == '__main__':
    nums1 = [-1,0,0,3,3,3,0,0,0]
    nums2 = [1,2,2]
    print(merge_88(nums1, nums2))
