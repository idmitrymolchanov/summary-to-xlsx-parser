package ru.example.exception;

import java.io.IOException;

public class WriteWorkbookException extends RuntimeException {

  public WriteWorkbookException(String message) {
    super(String.format("невозможно создать файл по пути: %s", message));
  }
}
