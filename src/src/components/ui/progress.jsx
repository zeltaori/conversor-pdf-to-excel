import React from "react";
import * as ProgressPrimitive from "@radix-ui/react-progress";
import { cn } from "@/lib/utils";

const Progress = React.forwardRef(
  ({ className, value = 0, ...props }, ref) => {
    const normalizedValue = Math.min(
      100,
      Math.max(0, Number(value) || 0)
    );

    return (
      <ProgressPrimitive.Root
        ref={ref}
        className={cn(
          "relative h-2 w-full overflow-hidden rounded-full bg-slate-200",
          className
        )}
        {...props}
      >
        <ProgressPrimitive.Indicator
          className="h-full w-full flex-1 bg-emerald-600 transition-all"
          style={{
            transform: `translateX(-${100 - normalizedValue}%)`,
          }}
        />
      </ProgressPrimitive.Root>
    );
  }
);

Progress.displayName = "Progress";

export { Progress };
