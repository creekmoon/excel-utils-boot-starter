package cn.creekmoon.excelUtils.example.config.swagger;

import lombok.Data;
import org.springframework.boot.autoconfigure.condition.ConditionalOnExpression;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

/**
 * @ClassName SwaggerProperties
 * @Description 加载swagger的配置文件
 */
@Configuration
@ConfigurationProperties(prefix = "swagger-config")
@ConditionalOnExpression("${swagger-config.enable:false}")
@Data
public class SwaggerProperties {
    private boolean enable;
    private String protocol;
    private String basePackage;
    private String title;
    private String version;
    private String description;
}
