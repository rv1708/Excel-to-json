import org.apache.http.client.HttpClient;
import org.apache.http.conn.ssl.SSLConnectionSocketFactory;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.ssl.SSLContextBuilder;
import org.springframework.http.client.HttpComponentsClientHttpRequestFactory;
import org.springframework.web.client.RestTemplate;

import javax.net.ssl.SSLContext;
import java.io.File;
import java.security.KeyStore;

public class SecureRestClient {

    public static void main(String[] args) throws Exception {

        // Path to keystore and truststore
        String keystorePath = "/path/to/your/keystore.jks";
        String truststorePath = "/path/to/your/truststore.jks";
        String keystorePassword = "your-keystore-password";
        String truststorePassword = "your-truststore-password";

        // Load Keystore
        KeyStore keystore = KeyStore.getInstance(KeyStore.getDefaultType());
        keystore.load(new java.io.FileInputStream(new File(keystorePath)), keystorePassword.toCharArray());

        // Load Truststore
        KeyStore truststore = KeyStore.getInstance(KeyStore.getDefaultType());
        truststore.load(new java.io.FileInputStream(new File(truststorePath)), truststorePassword.toCharArray());

        // Create SSLContext with both keystore and truststore
        SSLContext sslContext = SSLContextBuilder.create()
                .loadKeyMaterial(keystore, keystorePassword.toCharArray())
                .loadTrustMaterial(truststore, null)
                .build();

        // Create SSL socket factory
        SSLConnectionSocketFactory socketFactory = new SSLConnectionSocketFactory(sslContext);

        // Create HttpClient with the custom SSL socket factory
        HttpClient httpClient = HttpClients.custom()
                .setSSLSocketFactory(socketFactory)
                .build();

        // Create RestTemplate using HttpComponentsClientHttpRequestFactory
        HttpComponentsClientHttpRequestFactory factory = new HttpComponentsClientHttpRequestFactory(httpClient);
        RestTemplate restTemplate = new RestTemplate(factory);

        // Call API using RestTemplate
        String apiUrl = "https://api.example.com/secure-endpoint";
        String response = restTemplate.getForObject(apiUrl, String.class);
        System.out.println("API Response: " + response);
    }
}
