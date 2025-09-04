import { render, screen } from '@testing-library/react';
import App from './App';

test('renders Invoice Generator header', () => {
  render(<App />);
  const headerElement = screen.getByText(/Invoice Generator/i);
  expect(headerElement).toBeInTheDocument();
});
