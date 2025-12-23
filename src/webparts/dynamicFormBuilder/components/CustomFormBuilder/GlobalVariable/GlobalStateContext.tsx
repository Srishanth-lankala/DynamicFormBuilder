import React, { createContext, useState, useContext, ReactNode, useEffect } from 'react';

interface IGlobalState {
  globalVariable: string;
  setGlobalVariable: React.Dispatch<React.SetStateAction<string>>;
}

const GlobalStateContext = createContext<IGlobalState | undefined>(undefined);

export const GlobalStateProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  
  const savedGlobalVariable = localStorage.getItem('globalVariable');
  
  
  const [globalVariable, setGlobalVariable] = useState<string>(savedGlobalVariable || 'Initial Value');

  
  useEffect(() => {
    localStorage.setItem('globalVariable', globalVariable);
  }, [globalVariable]);  

  return (
    <GlobalStateContext.Provider value={{ globalVariable, setGlobalVariable }}>
      {children}
    </GlobalStateContext.Provider>
  );
};

export const useGlobalState = (): IGlobalState => {
  const context = useContext(GlobalStateContext);
  if (!context) {
    throw new Error('useGlobalState must be used within a GlobalStateProvider');
  }
  return context;
};
