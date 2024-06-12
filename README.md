# -
实现五年数据的数据分析

public class test35 {
    // main入口由OJ平台调用
    public static void main(String[] args) {
        Scanner cin = new Scanner(System.in, StandardCharsets.UTF_8.name());
        String information = cin.nextLine().trim();
        cin.close();
        int result = getMatchResult(information);
        System.out.println(result);
    }

    private static int getMatchResult(String information) {
        String[] exp = information.split("\\$");
        // 存储最终需要计算的$表达式的参数
        List<Integer> expFinal = new ArrayList<>();
        for (int i = 0; i < exp.length; i++) {
            if (exp[i].contains("@")) {
                String[] primExp = exp[i].split("@");
                expFinal.add(Arrays.stream(primExp).map(Integer::new).reduce((a, b) -> 2 * a + b + 3).orElse(0));
            } else {
                expFinal.add(Integer.parseInt(exp[i]));
            }
        }
        System.out.print(expFinal);
        return expFinal.stream().reduce((a, b) -> 3 * a + 2 * b + 1).orElse(0);
    }

