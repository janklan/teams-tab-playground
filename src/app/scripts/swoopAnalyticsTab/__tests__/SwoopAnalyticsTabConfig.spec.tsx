import * as React from "react";
import { shallow } from "enzyme";
import toJson from "enzyme-to-json";

import { SwoopAnalyticsTabConfig } from "../SwoopAnalyticsTabConfig";

describe("SwoopAnalyticsTabConfig Component", () => {
    // Snapshot Test Sample
    it("should match the snapshot", () => {
        const wrapper = shallow(<SwoopAnalyticsTabConfig />);
        expect(toJson(wrapper)).toMatchSnapshot();
    });

    // Component Test Sample
    it("should render the tab", () => {
        const component = shallow(<SwoopAnalyticsTabConfig />);
        const divResult = component.containsMatchingElement(<div>Configure your tab</div>);

        expect(divResult).toBeTruthy();
    });
});
