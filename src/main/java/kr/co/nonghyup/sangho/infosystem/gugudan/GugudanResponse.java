package kr.co.nonghyup.sangho.infosystem.gugudan;

import java.util.List;

public record GugudanResponse(
        int dan,
        List<Row> rows
) {
    public record Row(int multiplicand, int multiplier, int result) {
    }
}
