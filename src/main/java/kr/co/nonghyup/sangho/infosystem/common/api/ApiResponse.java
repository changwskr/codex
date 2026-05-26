package kr.co.nonghyup.sangho.infosystem.common.api;

import java.time.LocalDateTime;

public record ApiResponse<T>(
        String code,
        String message,
        T data,
        LocalDateTime timestamp
) {

    public static <T> ApiResponse<T> ok(T data) {
        return new ApiResponse<>("SUCCESS", "정상 처리되었습니다.", data, LocalDateTime.now());
    }
}
