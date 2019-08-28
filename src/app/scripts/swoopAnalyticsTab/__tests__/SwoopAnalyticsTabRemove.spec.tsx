import * as React from "react";
import { shallow } from "enzyme";
import toJson from "enzyme-to-json";

import { SwoopAnalyticsTabRemove } from "../SwoopAnalyticsTabRemove";

describe("SwoopAnalyticsTabRemove Component", () => {
    // Snapshot Test Sample
    it("should match the snapshot", () => {
        const wrapper = shallow(<SwoopAnalyticsTabRemove />);
        expect(toJson(wrapper)).toMatchSnapshot();
    });

    // Component Test Sample
    it("should render the tab", () => {
        const component = shallow(<SwoopAnalyticsTabRemove />);
        const divResult = component.containsMatchingElement(<div>You"re about to remove your tab...</div>);

        expect(divResult).toBeTruthy();
    });
});
