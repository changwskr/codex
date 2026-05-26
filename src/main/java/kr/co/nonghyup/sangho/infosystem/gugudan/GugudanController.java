package kr.co.nonghyup.sangho.infosystem.gugudan;

import jakarta.validation.constraints.Max;
import jakarta.validation.constraints.Min;
import kr.co.nonghyup.sangho.infosystem.common.api.ApiResponse;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/v1/gugudan")
public class GugudanController {

    private final GugudanService gugudanService;

    public GugudanController(GugudanService gugudanService) {
        this.gugudanService = gugudanService;
    }

    @GetMapping("/{dan}")
    public ApiResponse<GugudanResponse> get(
            @PathVariable @Min(1) @Max(9) int dan
    ) {
        return ApiResponse.ok(gugudanService.compute(dan));
    }
}
