import React from "react";
import { useNavigate } from "react-router-dom";

// Higher-order component to inject navigate function into class component
const withNavigate = (Component: React.ComponentType<any>) => {
  return function WrapperComponent(props: any) {
    const navigate = useNavigate();
    return <Component {...props} navigate={navigate} />
  };
};

export default withNavigate;


// USE THIS NAVIGATE PROPS TO NAVIGATE IN ANY PAGE 