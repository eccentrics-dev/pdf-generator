package src.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Template01Dto {
    private String name;
    private List<Route> routes;

    @Data
    public static class Route {
        private String name;
        private String id;
        private String status;

        public Route(String name, String id, String status) {
            this.name = name;
            this.id = id;
            this.status = status;
        }
    }
}
