package kr.co.nonghyup.sangho.infosystem.gugudan;

import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

@Service
public class GugudanService {

    private static final int MIN_DAN = 1;
    private static final int MAX_DAN = 9;
    private static final int MULTIPLIER_COUNT = 9;

    public GugudanResponse compute(int dan) {
        if (dan < MIN_DAN || dan > MAX_DAN) {
            throw new IllegalArgumentException(
                    "단(dan)은 %d 이상 %d 이하의 정수여야 합니다. 입력값: %d".formatted(MIN_DAN, MAX_DAN, dan)
            );
        }

        List<GugudanResponse.Row> rows = new ArrayList<>(MULTIPLIER_COUNT);
        for (int i = 1; i <= MULTIPLIER_COUNT; i++) {
            rows.add(new GugudanResponse.Row(dan, i, dan * i));
        }
        return new GugudanResponse(dan, rows);
    }
}
