package kr.co.nonghyup.sangho.infosystem.common.health;

import kr.co.nonghyup.sangho.infosystem.common.api.ApiResponse;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/v1/system")
public class SystemHealthController {

    @GetMapping("/health")
    public ApiResponse<SystemHealthResponse> health() {
        return ApiResponse.ok(new SystemHealthResponse("UP", "농협 상호금융 정보계 API"));
    }
}
