@regressionMarch 
Feature: Add Tariff Plan in demo site

Background:
        And User click on Add tariff plan button
        
@sanity1
  Scenario: Add Tariff Plan with hard coded datas
   //Given user line Hook class using
    When User Enters all the field
    And User click on  tariff plan submit button
    Then User should be displayed the congratulation you add tariff plan

@smoke
  Scenario: Add Tariff Plan with hard coded datas
   // Given User launches demo telecom site in Add Tariff plan
   // And User click on Add tariff plan button---- Background using
    When User Enters all the field with two dim list
      | 100 | 200 |  300 | 400 |  1 |  2 |  3 |
      | 200 | 400 |  600 | 800 |  2 |  4 |  6 |
      | 300 | 900 | 1200 | 300 |  3 |  9 |  9 |
      | 111 | 222 |  333 | 444 | 11 | 22 | 33 |
      | 888 | 444 |  666 | 222 |  8 |  6 |  4 |
    And User click on  tariff plan submit button
    Then User should be displayed the congratulation you add tariff plan

@thara
  Scenario: Add Tariff Plan with hard coded datas
   // Given User launches demo telecom site in Add Tariff plan
   // And User click on Add tariff plan button
    When User Enters all the field with two dim map
      | MR  | FLM | FIM  | FSP | LPMC | IPMC | SPC |
      | 100 | 200 |  300 | 400 |    1 |    2 |   3 |
      | 200 | 400 |  600 | 800 |    2 |    4 |   6 |
      | 300 | 900 | 1200 | 300 |    3 |    9 |   9 |
      | 111 | 222 |  333 | 444 |   11 |   22 |  33 |
      | 888 | 444 |  666 | 222 |    8 |    6 |   4 |
    And User click on  tariff plan submit button
    Then User should be displayed the congratulation you add tariff plan
 
 @sanity1
   Scenario Outline: 
   // Given User launches demo telecom site in Add Tariff plan
   // And User click on Add tariff plan button
    When User Enters all the field  "<MR>","<FLM>","<FIM>","<FSP>","<LPMC>","<IPMC>","<SPC>"
    And User click on  tariff plan submit button
    Then User should be displayed the congratulation you add tariff plan
    
    Examples:
      | MR  | FLM | FIM  | FSP | LPMC | IPMC | SPC |
      | 100 | 200 |  300 | 400 |    1 |    2 |   3 |
      | 200 | 400 |  600 | 800 |    2 |    4 |   6 |
      | 300 | 900 | 1200 | 300 |    3 |    9 |   9 |
      | 111 | 222 |  333 | 444 |   11 |   22 |  33 |
      | 888 | 444 |  666 | 222 |    8 |    6 |   4 |
      