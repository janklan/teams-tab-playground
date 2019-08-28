import * as React from "react";
import { shallow, mount } from "enzyme";
import toJson from "enzyme-to-json";

import { SwoopAnalyticsTab } from "../SwoopAnalyticsTab";

describe("SwoopAnalyticsTab Component", () => {
    // Snapshot Test Sample
    it("should match the snapshot", () => {
        const wrapper = shallow(<SwoopAnalyticsTab />);
        expect(toJson(wrapper)).toMatchSnapshot();
    });

    // Component Test Sample
    it("should render the tab", () => {
        const component = shallow(<SwoopAnalyticsTab />);
        const divResult = component.containsMatchingElement(<div>This is your tab</div>);

        expect(divResult).toBeTruthy();
    });

    // Mocking Sample
    it("should show alert on button click", () => {
        window.alert = jest.fn();

        const component = mount(<SwoopAnalyticsTab />);
        const button = component.find("button");
        button.simulate("click");

        expect(window.alert).toHaveBeenCalledWith("It worked!");

        component.unmount();
    });
});
