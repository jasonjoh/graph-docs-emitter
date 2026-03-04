import "../dist/src/index.js";

using TypeSpec.Reflection;

/**
 * Specify a custom example request body for an API method.
 * The value should be an object literal with property values
 * matching the expected request body format.
 */
extern dec exampleRequest(target: Operation, value: valueof unknown);

/**
 * Specify a custom example response body for an API method.
 * The value should be an object literal with property values
 * matching the expected response body format.
 */
extern dec exampleResponse(target: Operation, value: valueof unknown);
