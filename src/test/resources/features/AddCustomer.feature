@regressionFeb
Feature: Add Customer in demo site

@sanity1
  Scenario: Add Customer with hard coded datas
  
    And User click on add customer button
    When User enters all the fields
    And User click on submit button
    Then User should be displayed cutomer ID is generated or no
@smoke1 @thara
  Scenario: Add Customer with hard coded datas
    //Given User launches demo telecom site
    And User click on add customer button
    When User enters all the fields with one dimensional
      | gayu | rajesh | gayu11@gmail.com | thanjavur | 84849384 |
    And User click on submit button
    Then User should be displayed cutomer ID is generated or no
@thara
  Scenario: Add Customer with hard coded datas
   // Given User launches demo telecom site
    And User click on add customer button
    When User enters all the fields with one dimensional map
      | Fname   | poonga         |
      | Lname   | mani           |
      | Email   | park@gmail.com |
      | Address | pudukkottai    |
      | Phno    |       68667761 |
    And User click on submit button
    Then User should be displayed cutomer ID is generated or no
