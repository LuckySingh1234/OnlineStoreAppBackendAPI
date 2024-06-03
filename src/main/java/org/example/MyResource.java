package org.example;

import com.google.gson.Gson;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import jakarta.ws.rs.Consumes;
import jakarta.ws.rs.GET;
import jakarta.ws.rs.POST;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.QueryParam;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;
import org.json.JSONObject;

import java.util.List;

/**
 * Root resource (exposed at "myresource" path)
 */
@Path("myresource")
public class MyResource {
    @POST
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    @Path("managerLogin")
    public Response managerLogin(String userData) {
        JSONObject json = new JSONObject(userData);
        String email = json.getString("email");
        String password = json.getString("password");

        User user = new User(email, password);
        Manager signedInManager = Manager.login(user);
        JSONObject responseJson = new JSONObject();
        if (signedInManager != null) {
            responseJson.put("signedInManager", signedInManager.getEmail());
        }
        return Response.ok(responseJson.toString(), MediaType.APPLICATION_JSON).build();
    }

    @POST
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    @Path("getManagerActions")
    public Response fetchManagerActions(String reqData) {
        JSONObject json = new JSONObject(reqData);
        String email = json.getString("email");

        String managerActions = Manager.fetchManagerActions(email);
        JSONObject responseJson = new JSONObject();
        if (managerActions != null) {
            responseJson.put("managerActions", managerActions);
        }
        return Response.ok(responseJson.toString(), MediaType.APPLICATION_JSON).build();
    }

    @GET
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    @Path("getProductsByCategory")
    public Response fetchProductsByCategory(@QueryParam("category") String category) {
        List<Product> products = Product.fetchProductsByCategory(category);
        Gson gson = new Gson();
        JsonElement productsJson = gson.toJsonTree(products);
        return Response.ok(productsJson.toString(), MediaType.APPLICATION_JSON).build();
    }

    @POST
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    @Path("getProductById")
    public Response fetchProductById(String reqData) {
        JSONObject json = new JSONObject(reqData);
        String productId = json.getString("productId");
        Product product = Product.fetchProductById(productId);
        Gson gson = new Gson();
        JsonElement productJsonElement = gson.toJsonTree(product);
        if (productJsonElement.isJsonObject()) {
            JsonObject productJsonObject = productJsonElement.getAsJsonObject();
            productJsonObject.addProperty("success", "true");
            return Response.ok(gson.toJson(productJsonElement), MediaType.APPLICATION_JSON).build();
        } else {
            JsonObject resp = new JsonObject();
            resp.addProperty("error", "true");
            resp.addProperty("errorMessage", "Couldn't fetch product");
            return Response.ok(gson.toJson(resp), MediaType.APPLICATION_JSON).build();
        }
    }

    @POST
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    @Path("addProduct")
    public Response addProduct(String reqData) {
        JSONObject json = new JSONObject(reqData);
        String productId = json.getString("productId");
        String name = json.getString("name");
        String price = json.getString("price");
        String stockQuantity = json.getString("stockQuantity");
        String category = json.getString("category");
        String imageUrl = json.getString("imageUrl");
        String productAddedResponse = Product.addProduct(productId, name, price, stockQuantity, category, imageUrl);
        JSONObject responseJson = new JSONObject();
        if (productAddedResponse.equals("true")) {
            responseJson.put("success", "true");
            return Response.ok(responseJson.toString(), MediaType.APPLICATION_JSON).build();
        } else {
            responseJson.put("failure", "true");
            responseJson.put("errorMessage", productAddedResponse);
            return Response.ok(responseJson.toString(), MediaType.APPLICATION_JSON).build();
        }
    }
}
