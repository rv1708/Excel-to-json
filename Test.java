import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.ResponseEntity;
import org.springframework.web.client.RestTemplate;

public class RestClient {

    private static final String API_URL = "https://api.example.com/resource";

    public static void main(String[] args) {
        RestTemplate restTemplate = new RestTemplate();

        // Example: GET request
        String response = restTemplate.getForObject(API_URL, String.class);
        System.out.println("GET Response: " + response);

        // Example: POST request
        String requestBody = "{ \"key\": \"value\" }";
        HttpHeaders headers = new HttpHeaders();
        headers.set("Content-Type", "application/json");
        HttpEntity<String> request = new HttpEntity<>(requestBody, headers);

        ResponseEntity<String> postResponse = restTemplate.postForEntity(API_URL, request, String.class);
        System.out.println("POST Response: " + postResponse.getBody());

        // Example: PUT request
        restTemplate.put(API_URL, request);

        // Example: DELETE request
        restTemplate.delete(API_URL);
    }
}
