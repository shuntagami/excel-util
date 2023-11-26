import { APP_ENV } from "../../constant/constant";
import { BlueprintAPIClientService } from "./blueprintAPIClientService";
import { MockBlueprintAPIClientService } from "./mockblueprintAPIClientService";

let blueprintAPIClient: BlueprintAPIClientService

if (APP_ENV === "local") {
  blueprintAPIClient = new MockBlueprintAPIClientService
} else {
  blueprintAPIClient = new BlueprintAPIClientService
}

export { blueprintAPIClient };
