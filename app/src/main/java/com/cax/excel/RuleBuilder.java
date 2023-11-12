/*
 * RuleBuilder.java
 *
 * Description:
 * This Java class is a builder for constructing rules used in the Extractor class.
 * It creates rules in the form of JSONObjects with specific properties indicating directions and list indicators.
 *
 * Author: CodeArtisanX
 * Date: November 11, 2023
 */

package com.cax.excel;

import com.alibaba.fastjson.JSONObject;

/**
 * RuleBuilder class for constructing rules used in ExcelFieldValueExtractor.
 */
public class RuleBuilder {
    private final JSONObject rule;

    /**
     * Constructs a RuleBuilder with the specified field name.
     *
     * @param name The name of the field for which the rule is being constructed.
     */
    public RuleBuilder(String name) {
        this.rule = new JSONObject();
        rule.put("name", name);
        rule.put("right", false);
        rule.put("left", false);
        rule.put("up", false);
        rule.put("down", false);
        rule.put("list", false);
    }

    /**
     * Specifies a rule for moving to the right in the Excel sheet.
     *
     * @return The RuleBuilder instance for method chaining.
     */
    public RuleBuilder right() {
        rule.put("right", true);
        return this;
    }

    /**
     * Specifies a rule for moving to the left in the Excel sheet.
     *
     * @return The RuleBuilder instance for method chaining.
     */
    public RuleBuilder left() {
        rule.put("left", true);
        return this;
    }

    /**
     * Specifies a rule for moving upward in the Excel sheet.
     *
     * @return The RuleBuilder instance for method chaining.
     */
    public RuleBuilder up() {
        rule.put("up", true);
        return this;
    }

    /**
     * Specifies a rule for moving downward in the Excel sheet.
     *
     * @return The RuleBuilder instance for method chaining.
     */
    public RuleBuilder down() {
        rule.put("down", true);
        return this;
    }

    /**
     * Specifies a rule indicating that the field value is a list.
     *
     * @return The RuleBuilder instance for method chaining.
     */
    public RuleBuilder list() {
        rule.put("list", true);
        return this;
    }

    /**
     * Builds and returns the constructed rule as a JSONObject.
     *
     * @return The constructed rule as a JSONObject.
     */
    public JSONObject build() {
        return rule;
    }
}
