package kr.co.nonghyup.sangho.infosystem.common.health;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.test.web.servlet.MockMvc;

import static org.hamcrest.Matchers.is;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.jsonPath;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

@WebMvcTest(SystemHealthController.class)
class SystemHealthControllerTest {

    @Autowired
    private MockMvc mockMvc;

    @Test
    void healthReturnsUpStatus() throws Exception {
        mockMvc.perform(get("/api/v1/system/health"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.code", is("SUCCESS")))
                .andExpect(jsonPath("$.data.status", is("UP")));
    }
}
