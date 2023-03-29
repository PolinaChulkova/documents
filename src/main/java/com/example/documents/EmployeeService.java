package com.example.documents;

import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

@Service
@RequiredArgsConstructor
public class EmployeeService {

    private final EmployeeRepo employeeRepo;

    public List<Employee> addEmployees() {
        Employee employee1 = new Employee("Polina", "Chulkova", "Java developer");
        Employee employee2 = new Employee("Olga", "Orlova", "Manager");
        Employee employee3 = new Employee("Svetlana", "Korkova", "Scala developer");
        Employee employee4 = new Employee("Alexander", "Orlov", "Designer");

        employeeRepo.save(employee1);
        employeeRepo.save(employee2);
        employeeRepo.save(employee3);
        employeeRepo.save(employee4);

        List<Employee> employees = new ArrayList<>();
        employees.add(employeeRepo.findByFirstName("Polina"));
        employees.add(employeeRepo.findByFirstName("Olga"));
        employees.add(employeeRepo.findByFirstName("Svetlana"));
        employees.add(employeeRepo.findByFirstName("Alexander"));

        return employees;
    }
}
