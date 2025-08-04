def backtracking(candidates, ans, level, idx, target, memo):
    if target in memo:
        return memo[target]
    if target == 0:
        ans.append(level[:])
        return True
    if target < 0:
        return False

    found_combination = False
    for i in range(idx, len(candidates)):
        # Skip duplicate numbers
        if i > idx and candidates[i] == candidates[i - 1]:
            continue
        level.append(candidates[i])
        if backtracking(candidates, ans, level, i + 1, target - candidates[i], memo):
            found_combination = True
        level.pop()

    memo[target] = found_combination
    return found_combination


def combination_sum(candidates, target):
    ans = []
    level = []
    candidates.sort()
    memo = {}
    backtracking(candidates, ans, level, 0, target, memo)
    return ans


def main():
    total = int(input("輸入乘載總重: "))
    nums = list(map(int, input("輸入各商品重量: ").split()))

    ans = combination_sum(nums, total)
    while not ans:
        total -= 1
        ans = combination_sum(nums, total)
    print(total)
    for combination in ans:
        print(' '.join(map(str, combination)))
    print("end")


if __name__ == "__main__":
    main()