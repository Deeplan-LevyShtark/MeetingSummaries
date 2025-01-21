import React, { ReactNode } from 'react';
import { CacheProvider } from '@emotion/react';
import createCache from '@emotion/cache';
import rtlPlugin from 'stylis-plugin-rtl';

// Define the cache creation logic
const rtlCache = createCache({
    key: 'mui-rtl',
    stylisPlugins: [rtlPlugin],
});

const ltrCache = createCache({
    key: 'mui-ltr',
});

// Wrapper Component
interface CacheProviderWrapperProps {
    isRtl: boolean;
    children: ReactNode;
}

export const CacheProviderWrapper: React.FC<CacheProviderWrapperProps> = ({ isRtl, children }) => {
    const selectedCache = isRtl ? rtlCache : ltrCache;

    return <CacheProvider value={selectedCache}>{children}</CacheProvider>;
};
